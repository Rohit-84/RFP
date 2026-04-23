import json
from logger_config import get_logger

# Import the strict, zero-temperature LLM from your existing backend
from chatbot_backend import classifier_llm

logger = get_logger(__name__)

def generate_title_from_history(messages: list) -> str:
    """
    Analyzes the first few messages of a chat history to generate a smart,
    context-aware title using Azure OpenAI.
    """
    try:
        # 1. Grab up to the first 6 messages to get enough context
        # (Ignores the long tail of the conversation to save tokens and speed)
        early_messages = messages[:6]
        
        # 2. Format the messages into a readable transcript for the AI
        transcript = ""
        for msg in early_messages:
            role = "User" if msg.get("role") == "user" else "AI"
            content = msg.get("content", "").strip()
            if content:
                transcript += f"{role}: {content}\n"
        
        # 3. If there is no real transcript yet, return a default
        if not transcript.strip():
            return "New Chat"

        # 4. Prompt the AI to summarize the transcript
        prompt = f"""
        You are a highly intelligent title-generation assistant.
        Read the following chat transcript and generate a short, professional, 3 to 5 word title that captures the primary topic or goal of the user.

        Rules:
        1. If the user is only exchanging pleasantries (e.g., "Hi", "Hello", "How are you"), return exactly: "General Conversation".
        2. If the user is asking for a presentation/deck, capture the specific topic (e.g., "Deck: Snowflake Migration", "Proposal: Staples").
        3. Do NOT use punctuation like periods or quotes around the title.
        4. Output ONLY the title, nothing else.

        Transcript:
        {transcript}
        """

        # 5. Call the LLM to get the title
        resp_obj = classifier_llm.invoke(prompt)
        
        # 6. Safely extract and clean the text
        raw_title = resp_obj.content if hasattr(resp_obj, "content") else str(resp_obj)
        clean_title = raw_title.strip().strip('"').strip("'")
        
        # 7. Return the AI-generated title, or a fallback if it came back empty
        return clean_title if clean_title else "New Chat"

    except Exception as e:
        logger.error(f"Failed to generate title from history: {e}")
        # Fail silently and gracefully so the app never crashes
        return "New Chat"