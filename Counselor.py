import asyncio
import pandas as pd
from autogen_agentchat.agents import AssistantAgent
from autogen_ext.models.openai import OpenAIChatCompletionClient

# ==========================================
# 1. DATABASE / TOOL SECTOR
# ==========================================
def query_college_cutoffs(exam: str, student_percentile: float, category: str) -> str:
    """
    Queries historical database matrices for Maharashtra engineering college cutoffs.
    
    Args:
        exam: The name of the exam used for counseling ('MHT-CET' or 'JEE-Main').
        student_percentile: The float value score achieved by the student (0-100).
        category: The caste/reservation bucket of the candidate.
    """
    try:
        # Mocking data framework covering target categories and institutes
        mock_data = [
            {"college_name": "COEP Pune", "branch": "Computer Eng", "exam_type": "MHT-CET", "category": "Open", "cutoff_percentile": 99.8},
            {"college_name": "COEP Pune", "branch": "Computer Eng", "exam_type": "MHT-CET", "category": "OBC", "cutoff_percentile": 99.2},
            {"college_name": "VJTI Mumbai", "branch": "IT", "exam_type": "MHT-CET", "category": "OBC", "cutoff_percentile": 98.9},
            {"college_name": "PICT Pune", "branch": "Computer Eng", "exam_type": "MHT-CET", "category": "OBC", "cutoff_percentile": 97.8},
            {"college_name": "MIT WPU Pune", "branch": "Computer Eng", "exam_type": "MHT-CET", "category": "OBC", "cutoff_percentile": 93.5},
            {"college_name": "VIT Pune", "branch": "AI & DS", "exam_type": "MHT-CET", "category": "OBC", "cutoff_percentile": 94.8},
            {"college_name": "MIT WPU Pune", "branch": "Mechanical Eng", "exam_type": "JEE-Main", "category": "All-India", "cutoff_percentile": 85.0},
        ]
        df = pd.DataFrame(mock_data)
        
        filtered_df = df[(df['exam_type'].str.upper() == exam.strip().upper()) & 
                         (df['category'].str.upper() == category.strip().upper())]
        
        if filtered_df.empty:
            return f"No historical cutoff rows matched for {exam} under category {category}."
            
        matches = filtered_df[(filtered_df['cutoff_percentile'] <= student_percentile)]
        
        if matches.empty:
            return f"Your score of {student_percentile} is below historical cutoffs for the available records in this category."
            
        return f"DATABASE SEARCH RESULTS:\n{matches.to_string(index=False)}"
    except Exception as e:
        return f"Database Query Interrupted: {str(e)}"

# ==========================================
# 2. INSTRUCTION-BASED AGENT CONFIGURATION
# ==========================================
async def main():
    # Model configuration for Groq API
    # Configure Groq using the OpenAI-compatible client

    # Configure Groq using the OpenAI-compatible client
    model_client = OpenAIChatCompletionClient(
        model="llama-3.3-70b-versatile",            # <-- Change this string right here!
        api_key="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",      # <-- Keep your actual key here!
        base_url="https://api.groq.com/openai/v1",  
        temperature=0.1,
        model_info={
            "vision": False,
            "function_calling": True,
            "json_output": True,
            "family": "unknown"
        }
    )

    # Systematic Operational Guidelines for the Agent
    counselor_instructions = """
    ROLE: You are an expert Maharashtra Engineering CAP (Centralized Admission Process) Counselor Agent.
    
    OPERATIONAL DIRECTIVES:
    1. CONVERSATION INITIALIZATION:
       - Greet the student professionally.
       - If the student has not provided all details, you MUST explicitly ask for their:
         a) Exam Taken (MHT-CET or JEE Main)
         b) Percentile Score
         c) Caste/Reservation Category (Open, OBC, SC, ST, EWS, TFWS)
       - Do not process or guess predictions until all three details are provided.

    2. DATA ACCURACY STANDARD:
       - You are forbidden from inventing or hallucinating college names or cutoffs.
       - Once the student provides their three core metrics, you MUST invoke the `query_college_cutoffs` tool immediately.
       - Base your recommendations exclusively on the results returned by the tool.

    3. INSTITUTION BOUNDARY ENFORCEMENT:
       - Clearly remind students that IITs (like IIT Bombay) and NITs (like VNIT Nagpur) require JEE Advanced/JoSAA counseling and do not accept MHT-CET. 
       - Focus strictly on Maharashtra CAP institutions (Autonomous, Private, and Government colleges like COEP, VJTI, MIT-WPU, PICT).

    4. RESPONSE STRUCTURE:
       - Present your findings clearly using bullet points.
       - Group recommendations by availability and suitability.
       - End your response by asking the student if they want to check details for another branch or exam.
    """

    counselor_agent = AssistantAgent(
        name="Maharashtra_College_Counselor",
        model_client=model_client,
        system_message=counselor_instructions,
        tools=[query_college_cutoffs]
    )

    print("\n--- Maharashtra CAP Admission Counselor Agent Activated ---")
    print("Type 'exit' to end the session.\n")
    
    # Send an initial welcome message from the agent manually to start the loop cleanly
    print("Agent: Hello! I am your Maharashtra Admission Counselor. To help you find the best engineering colleges (Government, Private, or Autonomous), please tell me your Exam Type (MHT-CET/JEE Main), your Percentile Score, and your Category.")

    # ==========================================
    # 3. INTERACTIVE LIVE CHAT LOOP
    # ==========================================
    while True:
        user_input = input("\nYou (Student): ")
        if user_input.strip().lower() == 'exit':
            print("Ending session. Good luck with your admissions!")
            break
            
        if not user_input.strip():
            continue

        # Execute the agent workflow dynamically based on live human entry
        response = await counselor_agent.run(task=user_input)
        
        # Output the agent's response to the terminal console
        print(f"\nAgent: {response.messages[-1].content}")

if __name__ == "__main__":
    asyncio.run(main())