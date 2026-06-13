async def call_groq(
    prompt: str,
    file_content: str = "",
    file_type: str = "",
    image_data: dict = None,
    session_id: str = None,
) -> dict:

    messages = [{"role": "system", "content": SYSTEM_PROMPT}]

    # Inject conversation history
    if session_id and session_id in conversation_memory:
        history = conversation_memory[session_id][-6:]  # last 3 exchanges
        messages.extend(history)

    # Build user message
    if image_data:
        messages.append({
            "role": "user",
            "content": [
                {
                    "type": "image_url",
                    "image_url": {"url": f"data:{image_data['mime']};base64,{image_data['b64']}"}
                },
                {
                    "type": "text",
                    "text": f"""Analyze this image carefully.
{"Task: " + prompt if prompt.strip() else "Understand this image and create the best Excel spreadsheet from it."}
Read all text, numbers, tables visible in the image.
Solve any math problems, complete any missing data.
Return only JSON."""
                }
            ]
        })
    elif file_content and not prompt.strip():
        messages.append({
            "role": "user",
            "content": f"""I uploaded a {file_type} file. Carefully read EVERY row and understand what this file is about, then create a perfect professional Excel output.

FILE CONTENT (process ALL rows, do not skip any):
{file_content[:15000]}

CRITICAL INSTRUCTIONS:
- Read the file and understand its purpose automatically
- If it contains math problems (like "123 + 456"), solve ALL of them and put integer answers
- If it contains student marks, add Grade, Total, Rank, Pass/Fail
- If it contains inventory, add Stock Status, Value, Reorder Alert
- If it contains names/data with blanks, fill them intelligently
- Copy EVERY SINGLE ROW without skipping — if 100 rows in file, output 100 rows
- Add professional formulas, charts, conditional formatting
- Add TOTAL/AVERAGE/SUMMARY rows at the bottom
- Return only JSON"""
        })
    elif file_content and prompt.strip():
        messages.append({
            "role": "user",
            "content": f"""Task: {prompt}

FILE TYPE: {file_type}
FILE CONTENT (use ALL rows, do not skip any):
{file_content[:15000]}

IMPORTANT:
- Use every single row from the file
- Complete any missing values or answers
- Do the task described above using this real data
- Add professional formatting and formulas
- Return only JSON"""
        })
    else:
        messages.append({
            "role": "user",
            "content": f"""Create a professional Excel spreadsheet: "{prompt}"
- Real, meaningful data — minimum 15 rows
- Multiple sheets if it makes sense
- Professional formulas throughout
- Add chart if data is comparative or trending
- Add conditional formatting
- Summary row at bottom
- No placeholder values whatsoever
- Return only JSON"""
        })

    last_error = None
    
    # FIX: Use active production models on Groq
    TEXT_MODELS = [
        "llama-3.3-70b-versatile",
        "llama-3.3-70b-specdec",
        "llama3-70b-8192",
        "llama-3.1-8b-instant"
    ]
    IMAGE_MODELS = [
        "llama-3.2-11b-vision-preview",
        "llama-3.2-90b-vision-preview",
        "llama-3.3-70b-versatile"  # Fallback option if vision fails
    ]
    
    for attempt in range(4):
        try:
            temperature = [0.1, 0.2, 0.35, 0.5][attempt]
            if image_data:
                model = IMAGE_MODELS[min(attempt, len(IMAGE_MODELS)-1)]
                # Safe check if falling back to text-only model within image workflow
                is_vision_model = "vision" in model
                if not is_vision_model and isinstance(messages[-1]["content"], list):
                    # FIX: Keep structured dictionary array for user message block format
                    text_parts = [p["text"] for p in messages[-1]["content"] if p.get("type") == "text"]
                    messages[-1] = {
                        "role": "user", 
                        "content": [{"type": "text", "text": " ".join(text_parts)}]
                    }
            else:
                model = TEXT_MODELS[min(attempt, len(TEXT_MODELS)-1)]

            async with httpx.AsyncClient(timeout=60) as client:
                response = await client.post(
                    "https://api.groq.com/openai/v1/chat/completions",
                    headers={
                        "Authorization": f"Bearer {os.getenv('GROQ_API_KEY')}",
                        "Content-Type": "application/json"
                    },
                    json={
                        "model": model,
                        "max_tokens": 8000,
                        "temperature": temperature,
                        "messages": messages
                    }
                )

            if response.status_code == 429:
                logger.warning(f"Rate limit on {model}, trying next model...")
                last_error = f"Rate limit on {model}"
                await asyncio.sleep(1.5)  # Slightly longer cooldown backoff
                continue
            if response.status_code != 200:
                raise ValueError(f"Groq HTTP {response.status_code}: {response.text[:200]}")

            result = response.json()
            if "error" in result:
                raise ValueError(f"Groq error: {result['error']}")

            text = result["choices"][0]["message"]["content"].strip()
            text = re.sub(r"^```[a-zA-Z]*\n?", "", text)
            text = re.sub(r"\n?
```$", "", text)
            text = text.strip()

            # Find the outermost valid JSON object
            brace_start = text.find('{')
            if brace_start != -1:
                depth = 0
                brace_end = -1
                for i, ch in enumerate(text[brace_start:], brace_start):
                    if ch == '{':
                        depth += 1
                    elif ch == '}':
                        depth -= 1
                        if depth == 0:
                            brace_end = i + 1
                            break
                if brace_end != -1:
                    text = text[brace_start:brace_end]

            data = json.loads(text)

            valid, reason = validate_ai_response(data)
            if not valid:
                raise ValueError(f"Validation failed: {reason}")

            data = sanitize_ai_response(data)

            # Save to conversation memory
            if session_id:
                user_msg = {"role": "user", "content": prompt or f"[{file_type} file uploaded]"}
                assistant_msg = {"role": "assistant", "content": f"[Generated Excel with {len(data['sheets'])} sheet(s)]"}
                if session_id not in conversation_memory:
                    conversation_memory[session_id] = []
                conversation_memory[session_id].extend([user_msg, assistant_msg])
                conversation_memory[session_id] = conversation_memory[session_id][-20:]  # keep last 10 exchanges

            return data

        except json.JSONDecodeError as e:
            last_error = f"JSON parse error: {e}"
            logger.warning(f"Attempt {attempt+1} JSON error: {e}")
            continue
        except ValueError as e:
            last_error = str(e)
            logger.warning(f"Attempt {attempt+1} value error: {e}")
            continue
        except (httpx.TimeoutException, httpx.ReadTimeout, httpx.ConnectTimeout):
            last_error = "AI service timed out"
            logger.warning(f"Attempt {attempt+1} timeout on {model}")
            await asyncio.sleep(2)
            continue
        except Exception as e:
            last_error = str(e)
            logger.error(f"Attempt {attempt+1} unexpected error: {e}")
            continue

    raise ValueError(f"AI generation failed after 4 attempts. Last error: {last_error}")
            
