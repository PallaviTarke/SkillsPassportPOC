# Skill Passport Generator

A production-ready FastAPI application for curriculum skill extraction and student skill passport generation, powered by Google Gemini AI and deployed on Google Cloud Run.

## Features

- **AI-Powered Skill Extraction**: Extract skills from curriculum PDFs using Gemini 2.0 Flash
- **Bloom's Taxonomy Classification**: Automatic proficiency assessment based on course outcomes
- **Industry Skill Mapping**: LightCast integration for industry-standard skill taxonomy
- **McKinsey Framework**: Skills classified into McKinsey's 4-tier framework
- **Student Skill Passports**: Generate personalized skill passports from marksheets
- **Real-time Progress Tracking**: Server-sent events for extraction progress
- **Excel Export**: Professional Excel reports with conditional formatting
- **Cloud-Native**: Flexible authentication for local development and Cloud Run deployment

## Project Structure

```
skill_passport_final/
├── app/
│   ├── __init__.py
│   ├── main.py                      # FastAPI application entry point
│   ├── api/
│   │   ├── __init__.py
│   │   └── routes/
│   │       ├── __init__.py
│   │       ├── health.py            # Health check endpoints
│   │       ├── curriculum.py        # Curriculum management
│   │       ├── extraction.py        # Curriculum skill extraction
│   │       └── passport.py          # Student passport generation
│   ├── core/
│   │   ├── __init__.py
│   │   ├── config.py                # Configuration and clients
│   │   └── database.py              # MongoDB operations
│   ├── models/
│   │   ├── __init__.py
│   │   └── schemas.py               # Pydantic data models
│   ├── services/
│   │   ├── __init__.py
│   │   ├── extraction_service.py    # Curriculum extraction engine
│   │   └── passport_service.py      # Passport generation engine
│   └── utils/
│       └── __init__.py
├── .env                             # Local environment variables (not committed)
├── .env.example                     # Environment template
├── .dockerignore                    # Docker ignore patterns
├── Dockerfile                       # Container configuration
├── requirements.txt                 # Python dependencies
├── run.py                          # Local development runner
└── deploy-instructions.md           # Cloud Run deployment guide
```

## Quick Start

### Local Development Setup

1. **Clone and install dependencies**:
```bash
git clone <repository-url>
cd skill_passport_final
pip install -r requirements.txt
```

2. **Configure environment variables**:
```bash
# Copy the example file
cp .env.example .env

# Edit .env with your values
```

Required environment variables:
```env
# MongoDB
MONGO_URI=mongodb+srv://user:pass@cluster.mongodb.net/dbname
MONGO_DB_NAME=Aadhar_VC
LIGHTCAST_DB_NAME=Aadhar_VC
LIGHTCAST_COLLECTION=LightCast

# AI APIs
GEMINI_API_KEY=your-gemini-api-key
OPENAI_API_KEY=your-openai-api-key

# Google Document AI
DOCAI_PROJECT_ID=your-gcp-project-id
DOCAI_LOCATION=us
DOCAI_PROCESSOR_ID=your-processor-id

# Local development only
GOOGLE_APPLICATION_CREDENTIALS=./service-account.json
```

3. **Set up Google Cloud credentials** (for local development):
   - Download your service account JSON from Google Cloud Console
   - Save it as `service-account.json` in the project root
   - The file is already in `.gitignore` and won't be committed

## Running the Application

### Local Development (with auto-reload)
```bash
python run.py
```
The application will be available at `http://localhost:5000`

### Production Mode (Local)
```bash
uvicorn app.main:app --host 0.0.0.0 --port 8080
```

### Cloud Run Deployment

The application is Cloud Run-ready with automatic credential handling. See [deploy-instructions.md](deploy-instructions.md) for detailed deployment steps.

**Quick deploy command**:
```bash
gcloud run deploy skill-passport \
  --source . \
  --region us-central1 \
  --allow-unauthenticated \
  --env-vars-file .env.yaml \
  --memory 2Gi \
  --timeout 3600
```

**Authentication Handling**:
- **Local**: Uses `service-account.json` file
- **Cloud Run**: Automatically uses the service account attached to the Cloud Run service
- No code changes needed - the application detects the environment automatically

## API Documentation

### Interactive API Docs
- **Swagger UI**: `http://localhost:5000/docs`
- **ReDoc**: `http://localhost:5000/redoc`

### Core Endpoints

#### Health Check
```
GET  /                      # Application info and status
GET  /api/health            # System health with MongoDB connectivity
```

#### Curriculum Management
```
GET  /api/curriculum                    # Retrieve curriculum data with filters
GET  /api/curriculum/stats              # Statistics on stored curriculum
GET  /api/curriculum/options            # Available departments, years, semesters
```

#### Curriculum Extraction (Server-Sent Events)
```
POST /api/extract/start                 # Start extraction job
     Body: { pdf: File, department: str, year: str, semester: str, save_to_mongo: bool }
     Returns: { job_id: str }

GET  /api/extract/stream/{job_id}       # Real-time progress stream (SSE)
     Events: progress, done, error

GET  /api/extract/download/{job_id}     # Download Excel file
     Returns: Excel file with extracted skills
```

#### Student Skill Passport
```
POST /api/passport                      # Generate student passport
     Body: { pdf: File, student_name: str, roll_number: str,
             department: str, year: str, semester: str }
     Returns: { passport_id: str, ... }

GET  /api/passport/download/{passport_id}  # Download passport Excel
     Returns: Excel file with student's skill passport
```

### Example: Curriculum Extraction Flow

```javascript
// 1. Start extraction job
const formData = new FormData();
formData.append('pdf', pdfFile);
formData.append('department', 'Computer Science');
formData.append('year', 'MCA');
formData.append('semester', 'Semester 3');
formData.append('save_to_mongo', 'true');

const response = await fetch('/api/extract/start', {
    method: 'POST',
    body: formData
});
const { job_id } = await response.json();

// 2. Stream progress
const eventSource = new EventSource(`/api/extract/stream/${job_id}`);

eventSource.addEventListener('progress', (e) => {
    const data = JSON.parse(e.data);
    console.log(`${data.step}: ${data.msg} (${data.pct}%)`);
});

eventSource.addEventListener('done', (e) => {
    const data = JSON.parse(e.data);
    console.log(`Extracted ${data.count} skills`);
    eventSource.close();

    // 3. Download Excel
    window.location.href = `/api/extract/download/${job_id}`;
});

eventSource.addEventListener('error', (e) => {
    const data = JSON.parse(e.data);
    console.error('Error:', data.msg);
    eventSource.close();
});
```

## Technology Stack

### Backend Framework
- **FastAPI**: Modern async Python web framework
- **Uvicorn**: High-performance ASGI server
- **Pydantic**: Data validation and serialization

### AI & Machine Learning
- **Google Gemini 2.0 Flash**: Curriculum skill extraction and classification
- **OpenAI Embeddings**: Semantic skill matching (text-embedding-ada-002)
- **Google Document AI**: Advanced PDF text extraction with table detection

### Data & Storage
- **MongoDB**: Document database for curriculum and skill storage
- **LightCast Taxonomy**: Industry-standard skill taxonomy database
- **McKinsey Skills Framework**: 4-tier skill classification (Cognitive, Interpersonal, Self-Leadership, Digital)

### PDF Processing
- **PyMuPDF (fitz)**: Fast PDF text extraction
- **pdfplumber**: Table-aware PDF extraction with fallback
- **Google Document AI**: Cloud-based OCR and layout analysis

### Excel Generation
- **openpyxl**: Professional Excel reports with conditional formatting, filters, and multiple sheets

### Cloud & DevOps
- **Google Cloud Run**: Serverless container deployment
- **Docker**: Containerization
- **Google Cloud IAM**: Service account authentication

## How It Works

### Curriculum Skill Extraction Pipeline

1. **PDF Upload** → FastAPI receives curriculum PDF
2. **Text Extraction** → Document AI (chunked) → PyMuPDF → pdfplumber (fallback chain)
3. **AI Processing** → Gemini analyzes curriculum with EXTRACTION_PROMPT
   - Detects Bloom's taxonomy levels from course outcomes
   - Extracts granular skills (not vague umbrellas)
   - Identifies specific tools/technologies (raw_skill_keywords)
   - Assigns proficiency (Beginner/Intermediate/Advanced)
4. **Industry Mapping** → OpenAI embeddings + cosine similarity → top 15 LightCast candidates → Gemini picks best match
5. **Framework Classification** → McKinsey 4-tier framework classification
6. **Excel Generation** → Formatted workbook with summary and evidence sheets
7. **MongoDB Storage** → Skills saved for future passport generation

### Student Skill Passport Pipeline

1. **Marksheet Upload** → PDF containing student marks
2. **Curriculum Matching** → Fetch extracted curriculum from MongoDB
3. **Skill Aggregation** → Map student's subjects to skills
4. **Evidence-Based Proficiency** → Calculate based on:
   - Bloom's taxonomy level from curriculum
   - Student's actual marks
   - Contact hours and credits
5. **Passport Generation** → Professional Excel with skill summary and detailed evidence

## Configuration Details

### Flexible Authentication (Local + Cloud Run)

The application automatically detects the environment and uses appropriate credentials:

```python
# app/services/extraction_service.py
def _build_docai_client():
    if KEY_PATH.exists():
        # Local: Use service-account.json
        creds = service_account.Credentials.from_service_account_file(
            str(KEY_PATH), scopes=["https://www.googleapis.com/auth/cloud-platform"]
        )
        return documentai.DocumentProcessorServiceClient(credentials=creds)
    else:
        # Cloud Run: Use Application Default Credentials
        return documentai.DocumentProcessorServiceClient()
```

**Benefits**:
- ✅ No code changes between local and cloud environments
- ✅ No service account JSON in container images
- ✅ Secure credential management
- ✅ Works with any GCP service (Document AI, Vertex AI, etc.)

### PDF Extraction Strategy

The application uses a 3-tier fallback strategy for maximum reliability:

1. **Google Document AI** (Primary)
   - Chunked processing (15 pages per chunk)
   - Advanced table detection and layout analysis
   - Parallel chunk processing
   - Best for complex multi-column layouts

2. **PyMuPDF** (Fallback #1)
   - Fast text extraction
   - Block-based text ordering
   - Good for standard single-column PDFs

3. **pdfplumber** (Fallback #2)
   - Table-aware extraction
   - Handles scanned documents
   - Last resort for difficult PDFs

### AI Model Configuration

**Gemini 2.0 Flash**:
- Temperature: 0.05 (high consistency)
- Max tokens: 65,536
- Response format: JSON
- Retry logic: 4 attempts with exponential backoff

**OpenAI Embeddings**:
- Model: text-embedding-ada-002
- Batch processing for efficiency
- Used for semantic skill matching

### Excel Output Features

Generated Excel files include:

**Main Sheet**:
- 17 columns with skill details
- Conditional formatting for proficiency levels
- Color-coded Bloom's taxonomy
- Freeze panes and auto-filters
- Wrapped text in description columns

**Summary Sheet**:
- Department/Year/Semester metadata
- Total skills count
- LightCast mapping statistics
- McKinsey category distribution
- Proficiency breakdown
- Bloom's level distribution
- Skills per subject counts

**CO Evidence Sheet**:
- Detailed skill evidence
- Course outcome mappings
- Proficiency rationale
- Raw skill keywords

## Troubleshooting

### Common Issues

**MongoDB Connection Failed**
```bash
# Check connection string in .env
MONGO_URI=mongodb+srv://user:pass@cluster.mongodb.net/dbname?retryWrites=true

# Test connection
mongosh "mongodb+srv://user:pass@cluster.mongodb.net/dbname"
```

**Document AI Permission Denied**
```bash
# Local: Verify service account has Document AI API User role
# Cloud Run: Add role to Cloud Run service account
gcloud projects add-iam-policy-binding PROJECT_ID \
  --member="serviceAccount:SERVICE_ACCOUNT@PROJECT_ID.iam.gserviceaccount.com" \
  --role="roles/documentai.apiUser"
```

**Gemini API Rate Limits**
- The application includes automatic retry with exponential backoff
- Retries on: 503, 429, rate limit errors
- Max retries: 4 attempts

**PDF Extraction Returns No Text**
- Ensure PDF is not image-only (use Document AI for OCR)
- Check if PDF has text layer: `pdftotext file.pdf - | head`
- Document AI requires processor ID to be set in environment

## Performance & Scaling

### Local Development
- Single worker process
- Auto-reload enabled
- Suitable for testing and development

### Cloud Run Production
- **Memory**: 2GB recommended (handles large PDFs)
- **CPU**: 2 vCPU for parallel chunk processing
- **Timeout**: 3600s (1 hour) for large curriculum PDFs
- **Concurrency**: 80 requests per container (default)
- **Max instances**: 10 (adjust based on traffic)

### Optimization Tips
- Enable Cloud CDN for static assets
- Use Cloud Storage for large PDF storage
- Consider Cloud Tasks for async processing
- Implement request caching for curriculum data

## Security Considerations

- ✅ Service account JSON excluded from version control (`.gitignore`)
- ✅ Service account JSON excluded from container images (`.dockerignore`)
- ✅ Environment variables for sensitive data
- ✅ CORS middleware configured
- ✅ API key rotation recommended every 90 days
- ⚠️ Set `--no-allow-unauthenticated` for production Cloud Run deployment
- ⚠️ Use Google Secret Manager for production secrets

## License

Proprietary

## Support

For deployment assistance or issues, refer to:
- [deploy-instructions.md](deploy-instructions.md) - Detailed Cloud Run deployment guide
- API Documentation: `/docs` endpoint
- Application logs: `gcloud run services logs read skill-passport`
