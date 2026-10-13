"""
Default FastAPI application entry point for Vercel and ASGI runners.
Exposes the 'app' instance from web_server.
"""
from web_server import app

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)
