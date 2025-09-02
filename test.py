import os
import sys
import subprocess
import zipfile
import olefile
import re
import tempfile
import platform
from pathlib import Path
from fastapi import FastAPI, File, UploadFile, HTTPException, Request, Header
from fastapi.responses import JSONResponse
import PyPDF2
from typing import Optional
import ua_parser.user_agent_parser as ua_parser

app = FastAPI(title="Resume Parser API", description="API for counting pages in resume files", version="1.0.0")

def detect_server_os():
    """Detect the operating system where the server is running"""
    system = platform.system().lower()
    if 'windows' in system:
        return 'windows'
    elif 'darwin' in system:
        return 'macos'
    elif 'linux' in system:
        return 'linux'
    else:
        return 'unknown'

class ResumePageCounter:
    def __init__(self):
        self.server_os = detect_server_os()
        
    def count_pages(self, file_path: str, client_os: str = None) -> int:
        """
        Main function to count pages in DOC/DOCX/PDF files
        Returns the number of pages in the document
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"The file {file_path} does not exist")
        
        file_ext = os.path.splitext(file_path)[1].lower()
        
        if file_ext not in ['.doc', '.docx', '.pdf']:
            raise ValueError("Unsupported file format. Please provide a .doc, .docx, or .pdf file")
        
        # Handle PDF files
        if file_ext == '.pdf':
            return self._count_pdf_pages(file_path)
        
        # For Word documents, use server-side methods since client OS doesn't matter
        # for the actual processing (it happens on the server)
        try:
            if self.server_os == 'windows':
                page_count = self._count_pages_windows(file_path)
            elif self.server_os == 'macos':
                page_count = self._count_pages_macos(file_path)
            else:
                page_count = self._count_pages_cross_platform(file_path)
                
            if page_count > 0:
                return page_count
        except Exception as e:
            print(f"Primary method failed: {e}. Trying fallback methods...")
        
        # If OS-specific methods fail, try cross-platform methods
        try:
            page_count = self._count_pages_cross_platform(file_path)
            if page_count > 0:
                return page_count
        except Exception as e:
            print(f"Cross-platform method failed: {e}")
        
        # Final fallback: estimation based on file size
        return self._estimate_pages(file_path)
    
    def _count_pdf_pages(self, file_path: str) -> int:
        """Count pages in PDF files using PyPDF2"""
        try:
            with open(file_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                return len(pdf_reader.pages)
        except Exception as e:
            raise Exception(f"Failed to count PDF pages: {e}")
    
    def _count_pages_windows(self, file_path: str) -> int:
        """Count pages on Windows using Microsoft Word COM interface"""
        try:
            import win32com.client
            
            # Create Word application instance
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False
            
            # Open the document
            doc = word.Documents.Open(os.path.abspath(file_path))
            
            # Get accurate page count
            page_count = doc.ComputeStatistics(2)  # 2 = wdStatisticPages
            
            # Close documents and quit Word
            doc.Close(SaveChanges=False)
            word.Quit()
            
            return page_count
        except ImportError:
            print("pywin32 is not installed. Word COM interface unavailable.")
            raise Exception("Windows COM method not available on this server")
        except Exception as e:
            raise Exception(f"Windows COM method failed: {e}")
    
    def _count_pages_macos(self, file_path: str) -> int:
        """Count pages on macOS using AppleScript with Microsoft Word"""
        try:
            # AppleScript to get page count from Word
            applescript = f'''
            tell application "Microsoft Word"
                set myDoc to open "{os.path.abspath(file_path)}"
                set pageCount to count of pages of myDoc
                close myDoc saving no
                return pageCount
            end tell
            '''
            
            # Execute AppleScript
            process = subprocess.Popen(
                ['osascript', '-e', applescript],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True
            )
            stdout, stderr = process.communicate()
            
            if process.returncode == 0:
                return int(stdout.strip())
            else:
                raise Exception(f"AppleScript error: {stderr}")
                
        except Exception as e:
            raise Exception(f"macOS AppleScript method failed: {e}")
    
    def _count_pages_cross_platform(self, file_path: str) -> int:
        """Cross-platform method to extract page count from DOCX/DOC files"""
        file_ext = os.path.splitext(file_path)[1].lower()
        
        if file_ext == '.docx':
            return self._count_docx_pages(file_path)
        elif file_ext == '.doc':
            return self._count_doc_pages(file_path)
        else:
            return self._estimate_pages(file_path)
    
    def _count_docx_pages(self, file_path: str) -> int:
        """Extract page count from DOCX file (which is a ZIP archive)"""
        try:
            with zipfile.ZipFile(file_path) as docx:
                # Check if the document properties contain page count
                if 'docProps/app.xml' in docx.namelist():
                    with docx.open('docProps/app.xml') as app_xml:
                        content = app_xml.read().decode('utf-8')
                        # Look for Pages tag
                        match = re.search(r'<Pages>(\d+)</Pages>', content)
                        if match:
                            return int(match.group(1))
            
            # If page count not found in metadata, try to estimate from content
            return self._estimate_docx_pages(file_path)
        except:
            return self._estimate_pages(file_path)
    
    def _estimate_docx_pages(self, file_path: str) -> int:
        """Estimate page count for DOCX files by examining content"""
        try:
            from docx import Document
            
            doc = Document(file_path)
            total_content = 0
            
            # Count characters in paragraphs
            for paragraph in doc.paragraphs:
                total_content += len(paragraph.text)
            
            # Count characters in tables
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            total_content += len(paragraph.text)
            
            # Estimate pages based on content length
            # Adjust these values based on your typical resumes
            if total_content < 1500:
                return 1
            elif total_content < 3000:
                return 2
            else:
                return max(2, total_content // 1500)
        except ImportError:
            # python-docx not installed, fall back to file size estimation
            return self._estimate_pages(file_path)
        except:
            return self._estimate_pages(file_path)
    
    def _count_doc_pages(self, file_path: str) -> int:
        """Attempt to extract page count from binary DOC format"""
        # This is challenging without external libraries
        # For now, we'll just estimate based on file size
        return self._estimate_pages(file_path)
    
    def _estimate_pages(self, file_path: str) -> int:
        """Fallback method to estimate page count based on file size"""
        file_size = os.path.getsize(file_path)  # size in bytes
        
        # Rough estimation: 
        # - 40-50KB per page for DOCX 
        # - 30-40KB per page for DOC
        # - 50-100KB per page for PDF (varies based on content)
        file_ext = os.path.splitext(file_path)[1].lower()
        
        if file_ext == '.docx':
            return max(1, round(file_size / 45000))  # ~45KB per page
        elif file_ext == '.doc':
            return max(1, round(file_size / 35000))  # ~35KB per page
        else:  # .pdf
            return max(1, round(file_size / 75000))  # ~75KB per page

def detect_client_os(user_agent: str):
    """Attempt to detect client OS from User-Agent header"""
    if not user_agent:
        return 'unknown'
    
    try:
        parsed_string = ua_parser.Parse(user_agent)
        os_family = parsed_string['os']['family'].lower()
        
        if 'windows' in os_family:
            return 'windows'
        elif 'mac' in os_family or 'ios' in os_family:
            return 'macos'
        elif 'linux' in os_family or 'android' in os_family:
            return 'linux'
        else:
            return 'unknown'
    except:
        # Fallback simple detection
        user_agent_lower = user_agent.lower()
        if 'windows' in user_agent_lower:
            return 'windows'
        elif 'mac' in user_agent_lower:
            return 'macos'
        elif 'linux' in user_agent_lower:
            return 'linux'
        else:
            return 'unknown'

# FastAPI endpoints
@app.get("/")
async def root(request: Request, user_agent: Optional[str] = Header(None)):
    client_os = detect_client_os(user_agent)
    server_os = detect_server_os()
    
    return {
        "message": "Resume Parser API - Use /count-pages endpoint to count pages in resume files",
        "server_os": server_os,
        "client_os": client_os,
        "note": "File processing happens on the server, so server OS determines available methods"
    }

@app.post("/count-pages")
async def count_pages(
    file: UploadFile = File(...),
    user_agent: Optional[str] = Header(None)
):
    """
    Count pages in uploaded resume file (DOC, DOCX, or PDF)
    """
    # Validate file type
    if not file.filename:
        raise HTTPException(status_code=400, detail="No file provided")
    
    file_ext = Path(file.filename).suffix.lower()
    if file_ext not in ['.doc', '.docx', '.pdf']:
        raise HTTPException(
            status_code=400, 
            detail="Unsupported file format. Please upload a .doc, .docx, or .pdf file"
        )
    
    # Save uploaded file to temporary location
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=file_ext) as temp_file:
            content = await file.read()
            temp_file.write(content)
            temp_file_path = temp_file.name
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to process uploaded file: {str(e)}")
    
    # Count pages
    try:
        counter = ResumePageCounter()
        page_count = counter.count_pages(temp_file_path)
        
        # Detect client OS for information purposes
        client_os = detect_client_os(user_agent)
        server_os = detect_server_os()
        
        # Clean up temporary file
        try:
            os.unlink(temp_file_path)
        except:
            pass
            
        return JSONResponse(
            status_code=200,
            content={
                "filename": file.filename,
                "page_count": page_count,
                "file_type": file_ext[1:],  # Remove the dot
                "server_os": server_os,
                "client_os": client_os,
                "processing_note": f"File was processed on a {server_os} server"
            }
        )
    except Exception as e:
        # Clean up temporary file even if there's an error
        try:
            os.unlink(temp_file_path)
        except:
            pass
            
        raise HTTPException(status_code=500, detail=f"Failed to count pages: {str(e)}")

@app.get("/health")
async def health_check():
    """Health check endpoint"""
    server_os = detect_server_os()
    return {
        "status": "healthy", 
        "server_os": server_os,
        "server_platform": platform.platform(),
        "note": "This shows the OS where the server is running, not the client OS"
    }

@app.get("/debug")
async def debug_info(request: Request, user_agent: Optional[str] = Header(None)):
    """Debug endpoint to see server and client information"""
    client_os = detect_client_os(user_agent)
    server_os = detect_server_os()
    
    return {
        "server_os": server_os,
        "server_platform": platform.platform(),
        "client_os": client_os,
        "user_agent": user_agent,
        "client_host": request.client.host if request.client else "unknown",
        "note": "File processing uses server OS methods, not client OS methods"
    }

# Command line interface
def main():
    import argparse
    
    parser = argparse.ArgumentParser(description='Count pages in DOC/DOCX/PDF files')
    parser.add_argument('file_path', help='Path to the DOC, DOCX, or PDF file')
    args = parser.parse_args()
    
    counter = ResumePageCounter()
    
    try:
        page_count = counter.count_pages(args.file_path)
        server_os = detect_server_os()
        print(f"Server OS: {server_os}")
        print(f"The document has {page_count} page(s)")
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)

if __name__ == "__main__":
    # If run directly, use command line interface
    if len(sys.argv) > 1:
        main()
    else:
        # Import uvicorn to run the FastAPI app
        import uvicorn
        uvicorn.run(app, host="0.0.0.0", port=8000)
