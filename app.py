import asyncio
from playwright_stealth import Stealth
from playwright.async_api import async_playwright

from fastapi import FastAPI, Form
from fastapi.responses import FileResponse, HTMLResponse
from pydantic import BaseModel

class userinput(BaseModel):
    text: str

app = FastAPI()


@app.get("/")
async def root():
    return FileResponse("templates/index.html")

@app.post("/userinput", response_class=HTMLResponse)
async def open_page(userinput:userinput):
    async with Stealth().use_async(async_playwright()) as playwright:
        browser = await playwright.chromium.launch(headless=True)
        page = await browser.new_page(viewport={"width": 1920, "height": 1080})
        await page.goto(userinput.text)
        html_content = await page.content()
        return html_content

    

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)










""""
async def main():
    async with Stealth().use_async(async_playwright()) as playwright:
        browser = await playwright.chromium.launch(headless=False)
        page = await browser.new_page()
        await page.goto("https://x.com/elonmusk")
        await page.pause(); 

asyncio.run(main())
"""