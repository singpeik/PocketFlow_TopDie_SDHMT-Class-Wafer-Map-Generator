
<!--
TEMPLATE USAGE INSTRUCTIONS:
This file serves as a template for clearly defining your web application requirements.
When creating your own project, replace the example content below with your specific needs.

FOR FRONTEND REQUIREMENTS:
- Clearly describe the user interface components (forms, buttons, displays)
- Specify user interactions and workflows step-by-step
- Define the visual layout and user experience expectations
- Include any specific UI frameworks or libraries you prefer (React, Vue, vanilla JS, etc.)
- Describe responsive design requirements and browser compatibility

FOR BACKEND REQUIREMENTS:  
- Detail the API endpoints and their functionality
- Specify data processing workflows and business logic
- Define integration requirements (external APIs, databases, file systems)
- Include authentication, authorization, and security requirements
- Specify performance requirements and constraints
- Detail error handling and logging expectations

Be as specific as possible - this helps ensure the AI understands exactly what you want to build.
Replace this entire example below with your actual project requirements.

Please tell your agent that you are using the file as a way to guide the creation of the prompt that you use to create your design. Ask you agent to reformat it if needed to better guide the design. Ask it explicitly not to implement a design - ONLY to reformat this file in the best possible to guide the design itself. 

After you are done filling this out (with or without the help of the coding agent), please have the coding agent review the entire prompt template and ask you questions that it feels are necessary to create a better design for the app. Have it ask you the questions one at a time and have the questions in a multiple choice format. That should help you build a clearer picture of your app before you construct your design. 

-->

# Glossary Generator Web App - Requirements

I am trying to build a FastAPI web app with Swagger UI for documentation. My application takes in multiple user specified PDF documents, calls a user specified LLM endpoint with a user specified API key, and creates a glossary of unique terms from these documents and saves it as a PDF called glossary.pdf in a user specified location.

## Web App Route Description

The web app has the following routes:
1. `"/"` where the UI sits
2. `"/docs"` where the Swagger UI sits  
3. `"/health"` where the user can get a health update on the app

## UI Description

In the UI the user will:
1. Drag and drop documents into a file upload box
2. Provide a location of .env file containing the LLM endpoint, and the API key in a separate input box
3. Provide the output directory in which to store the glossary.pdf in a separate input box
4. Push an execute button to run the backend

Please design the appropriate frontend using the above description to support the necessary functions. Please use JS, HTML, TS etc. as needed. Please add jinja2 or any other suitable library to the requirements.txt here.

## Backend Description

I would like my app to have a python backend. When the execute button is pushed by the user in the UI:

1. The backend will ingest the documents inputted by the user in the UI described above
2. The backend will load the .env file from the location provided by the user in the UI described above and get the LLM endpoint and API key in there
3. It will use the input documents, LLM endpoint, API key, and use the call_llm.py function in the utils along with a system prompt that I will provide in a YAML config file in the utils directory to extract a glossary from the pdfs
4. Once it has the glossary it will print them out to a pdf called "glossary.pdf" to the user specified output directory from the UI

The pdf will look something like this:

```
Term1 - definition1 (pdf file name where the definition1 is located)
Term2 - definition2 (pdf file name where the definition2 is located)  
Term3 - definition3 (pdf file name where the definition3 is located)
... and so on.
```

If a term has multiple definitions it should have multiple rows.

## Processing Specifications

### Term Extraction Criteria
- Allow users to configure what types of terms to extract via UI checkboxes:
  - Technical and domain-specific terms
  - Acronyms and abbreviations 
  - Proper nouns and specialized concepts
- Users can select multiple term types for comprehensive extraction

### File Processing Limits
- Maximum 5 PDF files per batch processing session
- Maximum 10MB per individual PDF file (50MB total per session)
- Support standard PDF formats, handle text-based PDFs primarily

### Duplicate Term Handling
- When the same term appears in multiple documents with different definitions:
- Merge all definitions into single entries per term
- Format: "Term - Definition 1; Definition 2; Definition 3 (source1.pdf, source2.pdf, source3.pdf)"
- Combine source references for comprehensive tracking

### Output Format
- Generate PDF with terms sorted alphabetically (A-Z)
- Simple, clean format: "Term - Definition (source filenames)"
- Standard glossary layout for easy reference and lookup

## User Experience Requirements

### Progress Feedback
- Display progress bar with percentage completion (0-100%)
- Show current processing step (e.g., "Processing file 2 of 5...", "Extracting terms...", "Generating PDF...")
- Include estimated time remaining when possible
- Clear visual indication of processing status

### Error Handling Strategy
- Skip problematic files and continue processing remaining files
- Generate partial glossary from successfully processed documents
- Display summary at completion showing:
  - Successfully processed files
  - Skipped files with brief error reasons
  - Final glossary statistics (total terms extracted)

## Technical Requirements

### LLM Provider Support
- Support any OpenAI-compatible API endpoint (universal compatibility)
- Accept custom endpoint URLs and API keys via .env file
- Support for custom headers and authentication methods
- Compatible with local models (Ollama), cloud providers (OpenAI, Anthropic, etc.)

### Core Framework Requirements
- Use FastAPI for the backend web framework
- Use the existing `call_llm.py` utility function for LLM interactions
- Support user-configurable LLM endpoints and API keys via .env files
- Process multiple PDF documents for glossary extraction
- Generate formatted PDF output with term definitions and source references
- Include proper error handling and user feedback
- Provide Swagger/OpenAPI documentation
