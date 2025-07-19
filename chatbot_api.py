from dotenv import load_dotenv
import os

load_dotenv()

DATABRICKS_TOKEN = os.environ.get("DATABRICKS_TOKEN")

from openai import OpenAI
client = OpenAI(
    api_key=DATABRICKS_TOKEN,
    base_url="https://dbc-1bd7d2ac-5d2d.cloud.databricks.com/serving-endpoints"
)


def get_chat_response(messages):
    # For Databricks agents, we need to send conversation context properly
    # But avoid overwhelming the API with too much history
    
    # Take only the last 5 exchanges (10 messages) to maintain context while staying efficient
    recent_messages = messages[-10:] if len(messages) > 10 else messages
    
    # Format messages for Databricks agent
    valid_messages = []
    
    for m in recent_messages:
        # Only include messages with proper content
        if ("role" in m and "content" in m and 
            str(m["content"]).strip() and 
            m["role"] in ["user", "assistant"]):
            
            clean_message = {
                "role": m["role"],
                "content": str(m["content"]).strip()
            }
            valid_messages.append(clean_message)

    # Ensure we have at least the latest user message
    if not valid_messages and messages:
        last_message = messages[-1]
        if "role" in last_message and "content" in last_message:
            valid_messages = [{
                "role": last_message["role"],
                "content": str(last_message["content"]).strip()
            }]

    if not valid_messages:
        return "No valid messages to process."

    # Debug: Print the messages being sent
    print(f"Sending {len(valid_messages)} messages to Databricks:")
    for i, msg in enumerate(valid_messages):
        print(f"Message {i}: {msg['role']} - {msg['content'][:100]}...")

    try:
        # Try with a lower-level approach for Databricks
        import requests
        import json
        
        headers = {
            "Authorization": f"Bearer {DATABRICKS_TOKEN}",
            "Content-Type": "application/json"
        }
        
        payload = {
            "messages": valid_messages,
            "max_tokens": 2000,
            "temperature": 0.7
        }
        
        response = requests.post(
            "https://dbc-1bd7d2ac-5d2d.cloud.databricks.com/serving-endpoints/agents_workspace-default-quickstart_agent/invocations",
            headers=headers,
            json=payload,
            timeout=30
        )
        
        if response.status_code == 200:
            result = response.json()
            
            # Handle Databricks agent response format
            if 'messages' in result and result['messages']:
                messages = result['messages']
                
                # Look for tool calls (code execution)
                for msg in messages:
                    if msg.get('role') == 'assistant' and 'tool_calls' in msg:
                        # Extract the code from tool calls
                        for tool_call in msg['tool_calls']:
                            if tool_call.get('function', {}).get('name') == 'system__ai__python_exec':
                                import json
                                try:
                                    args = json.loads(tool_call['function']['arguments'])
                                    code = args.get('code', '')
                                    if code:
                                        return f"Here's a Python scientific calculator:\n\n```python\n{code}\n```"
                                except:
                                    pass
                    
                    # If it's a tool response with content, include that too
                    elif msg.get('role') == 'tool' and 'content' in msg:
                        continue  # Skip tool output for now
                    
                    # Regular message content
                    elif 'content' in msg and msg['content']:
                        content = msg['content']
                        # Don't return raw file content dumps
                        if "--- File Content" in content:
                            # Try to extract meaningful response after file content
                            parts = content.split("--- File Content")
                            if len(parts) > 1:
                                after_content = parts[1]
                                if "---" in after_content:
                                    response_parts = after_content.split("---", 1)
                                    if len(response_parts) > 1 and response_parts[1].strip():
                                        return response_parts[1].strip()
                            # If no meaningful content after, return the part before
                            if parts[0].strip():
                                return parts[0].strip()
                        return content
                
                # Fallback: return the last message content
                return messages[-1].get('content', 'No response content found')
            
            # Other response formats
            elif 'choices' in result and result['choices']:
                return result['choices'][0]['message']['content']
            elif 'content' in result:
                return result['content']
            else:
                return f"Unexpected response format: {result}"
        else:
            return f"HTTP Error {response.status_code}: {response.text}"
            
    except Exception as e:
        error_msg = f"Error: {str(e)}"
        return error_msg

