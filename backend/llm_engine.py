import os
from dotenv import load_dotenv
from langchain_groq import ChatGroq
import json
from langchain_core.prompts import ChatPromptTemplate

load_dotenv()

# ENGINE 1: Normal LLM (Text aur Keywords nikalne ke liye)
llm_text = ChatGroq(
    temperature=0.1, 
    model_name="llama-3.3-70b-versatile"
)

# ENGINE 2: Strict JSON LLM (Sirf Resume rewrite dictionary ke liye)
llm_json = ChatGroq(
    temperature=0.1, 
    model_name="llama-3.3-70b-versatile",
    model_kwargs={"response_format": {"type": "json_object"}}
)

def extract_keywords(jd_text):
    """JD se top keywords nikalne ke liye"""
    prompt = ChatPromptTemplate.from_messages([
        ("system", "Extract the top 8 most critical technical skills or keywords from this JD. Output ONLY a comma-separated list. No introduction, no extra words."),
        ("human", "{jd}")
    ])
    # Yahan hum Engine 1 (llm_text) use kar rahe hain
    response = (prompt | llm_text).invoke({"jd": jd_text})
    return response.content

def generate_tailored_resume_json(jd_text, relevant_chunks, original_resume_text):
    system_prompt = """
    YYou are a highly logical Expert ATS Resume Writer. Your job is to ENHANCE the resume, NOT rewrite history.
    
    STRICT RULES (READ CAREFULLY):
    1. NO FAKE FEATURES: DO NOT remove the user's existing work, features, or metrics. DO NOT invent new projects.
    2. SMART INTEGRATION: Find 3-4 experience/project bullet points. Inject JD keywords by logically connecting them to what the user ALREADY did. 
       (e.g., If they built a web app, and the JD needs 'AWS', change it to "Built a web app and deployed using AWS" - DO NOT replace the web app).
    3. SKILLS SECTION: Find the paragraph that lists the user's technical skills. APPEND the missing JD keywords to the end of that list. Do not delete their old skills.
    4. EXACT MATCH: The key in the JSON must be the EXACT WORD-FOR-WORD old text from the resume.
    
    Output ONLY a valid JSON object. No intro, no markdown blocks.
    Example:
    {{
        "Java, Python, C++": "Java, Python, C++, Docker, Kubernetes, AWS",
        "Developed a user authentication system.": "Developed a secure user authentication system, integrating OAuth 2.0 as per enterprise standards."
    }}
    """
    
    prompt = ChatPromptTemplate.from_messages([
        ("system", system_prompt),
        ("human", "Original Resume:\n{original}\n\nJD:\n{jd}\n\nContext:\n{context}")
    ])
    
    context = "\n".join([doc.page_content for doc in relevant_chunks])
    
    # Yahan hum Engine 2 (llm_json) use kar rahe hain
    chain = prompt | llm_json
    response = chain.invoke({"original": original_resume_text, "jd": jd_text, "context": context})
    
    try:
        clean_text = response.content.replace("```json", "").replace("```", "").strip()
        changes_dict = json.loads(clean_text)
        return changes_dict
    except Exception as e:
        print("JSON Parsing Error:", e)
        return {}