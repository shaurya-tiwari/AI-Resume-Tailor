from langchain_text_splitters import RecursiveCharacterTextSplitter
from langchain_community.vectorstores import FAISS
from langchain_community.embeddings import HuggingFaceEmbeddings

def create_vector_store(resume_text):
    # 1. Chunking: Text ko chhote, logical hisso mein todna
    text_splitter = RecursiveCharacterTextSplitter(chunk_size=500, chunk_overlap=50)
    chunks = text_splitter.split_text(resume_text)
    
    # 2. Embeddings: Text ka meaning samajhne ke liye 'all-MiniLM' model
    embeddings = HuggingFaceEmbeddings(model_name="all-MiniLM-L6-v2")
    
    # 3. Vector DB: Embeddings ko FAISS mein store karna
    vector_store = FAISS.from_texts(chunks, embeddings)
    
    return vector_store