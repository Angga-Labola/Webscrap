import asyncio
from fastapi import FastAPI, Form, Request
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates
from playwright.async_api import async_playwright
import os

app = FastAPI()
templates = Jinja2Templates(directory="templates")

@app.get("/", response_class=HTMLResponse)
async def read_form(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.post("/browse", response_class=HTMLResponse)
async def browse_link(request: Request, url: str = Form(...)):
    screenshot_path = "static/screenshot.png"
    page_title = ""
    page_content = ""
    error_message = None

    try:
        async with async_playwright() as p:
            # Launch the browser headlessly
            browser = await p.chromium.launch(headless=True)
            
            # Create a completely isolated, incognito browser context
            context = await browser.new_context()
            
            # Open a new page within the incognito context
            page = await context.new_page()
            
            # Navigate to the user-provided URL
            await page.goto(url, wait_until="networkidle", timeout=15000)
            
            # Extract data from the page
            page_title = await page.title()
            
            # Get clean inner text from the body element
            page_content = await page.locator("body").inner_text()
            
            # Truncate content for display readability
            page_content = page_content[:2000] + "..." if len(page_content) > 2000 else page_content

            # Clean up the session
            await context.close()
            await browser.close()
            
    except Exception as e:
        error_message = f"Failed to browse page: {str(e)}"

    return templates.TemplateResponse(
        "index.html", 
        {
            "request": request, 
            "url": url, 
            "title": page_title, 
            "content": page_content,
            "error": error_message
        }
    )

if __name__ == "__main__":
    import uvicorn

    port = int(os.getenv("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
