# SummarAIze
PDF to PowerPoint Summarizer A Streamlit app that converts PDFs into summarized PowerPoint presentations. Uses SpaCy for NLP summarization, KMeans for topic clustering, and python-pptx for slide creation. Customizable summary ratio, topics, and presentation details for efficient content extraction.


**PDF to PowerPoint Summarizer and Topic Extractor**  
This project is a powerful tool designed to convert lengthy PDF documents into concise PowerPoint presentations. Using NLP techniques and clustering algorithms, it efficiently summarizes content and organizes it into well-structured slides.  

### **Key Features**
✅ Extracts text from PDF files and summarizes it using SpaCy NLP.  
✅ Uses **KMeans clustering** with **TF-IDF vectorization** to identify key topics.  
✅ Generates PowerPoint presentations with organized slides, leveraging an "Organic" template.  
✅ Customizable options for summary ratio, number of topics, and presentation title/subtitle.  
✅ Built with **Streamlit** for an interactive and user-friendly interface.  

### **Tech Stack**
- **Python** (for data processing and logic)  
- **SpaCy** (for NLP text summarization)  
- **Scikit-learn** (for clustering and keyword extraction)  
- **PyPDF2** (for PDF text extraction)  
- **python-pptx** (for PowerPoint creation)  
- **Streamlit** (for web-based UI)  

### **How to Use**
1. Upload a PDF document.  
2. Choose the desired **summary ratio** and **number of topics**.  
3. Provide a **title** and **subtitle** for the presentation.  
4. Click **"Generate PowerPoint"** to download a structured presentation.  

This project is ideal for students, professionals, and researchers who need to extract key insights from lengthy documents efficiently.  
