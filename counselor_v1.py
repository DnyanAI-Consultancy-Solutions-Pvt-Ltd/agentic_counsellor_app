import os
import autogen
from autogen import AssistantAgent
from autogen.agentchat.contrib.retrieve_user_proxy_agent import RetrieveUserProxyAgent

# 1. LLM API Configurations (Using your Groq Key via OpenAI-Compatible Endpoint)
config_list = [
    {
        "model": "llama-3.3-70b-versatile",
        "api_key": "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",      # Groq API key!
        "base_url": "https://api.groq.com/openai/v1",
        "temperature": 0.1
    }
]

llm_config = {
    "config_list": config_list,
    "timeout": 600,
    "cache_seed": None  # Set to None to prevent local cache forcing static replies
}

# 2. Instantiate the Core Counselor Assistant
counselor_assistant = AssistantAgent(
    name="counselor_assistant",
    system_message="""You are an expert Maharashtra Engineering CAP Counselor Agent.
    Your objective is to answer college admission rules, eligibility guidelines, and seat allocations.
    Analyze the background context provided by the RAG proxy agent carefully to synthesize your answers.
    If you cannot find the answer within the context provided, state clearly that the rulebook does not contain it.""",
    llm_config=llm_config,
)

# 3. Instantiate the Native AutoGen 0.2 RetrieveUserProxyAgent
rag_user_proxy = RetrieveUserProxyAgent(
    name="rag_user_proxy",
    human_input_mode="ALWAYS", # Set to ALWAYS so it allows you to type interactively
    max_consecutive_auto_reply=3,
    retrieve_config={
        "task": "qa", # Standard Question-Answering Task Type
        
        # ────────── PLACE YOUR PDF LINKS/PATHS HERE ──────────
        "docs_path": [
            "2023SeatMatrix.pdf", 
            "2023ENGG_CAP1_CutOff.pdf"
        ],
        # ─────────────────────────────────────────────────────
        
        "chunk_token_size": 1000,
        "model": config_list[0]["model"],
        "vector_db": "chroma", # Utilizing local chromadb setup natively
        "collection_name": "maharashtra_admission_docs",
        "get_or_create": True, # Reuses database chunks if files don't change
    },
    code_execution_config=False,
)

# 4. Interactive Command-Line Console Loop
print("\n--- Maharashtra CAP Counselor (AutoGen 0.2 Native RAG) Activated ---")
print("Ask any rule-based, cutoff or criteria question below.")

while True:
    student_question = input("\nYou (Student): ")
    if student_question.strip().lower() == "exit":
        print("Ending evaluation session. Best of luck!")
        break
    if not student_question.strip():
        continue
        
    # Reset assistant state between different questions to avoid conversation mixing
    counselor_assistant.reset()
    
    # Fire up the native 0.2 RAG pipeline sequence
    rag_user_proxy.initiate_chat(
        counselor_assistant,
        message=rag_user_proxy.message_generator,
        problem=student_question
    )