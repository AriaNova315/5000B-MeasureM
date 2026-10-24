from contextlib import asynccontextmanager
from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse

from app.database import init_db
from app.routers import projects, export


@asynccontextmanager
async def lifespan(app: FastAPI):
    await init_db()
    yield


app = FastAPI(title="MeasureM", lifespan=lifespan, docs_url=None, redoc_url=None)

app.include_router(projects.router)
app.include_router(export.router)

app.mount("/static", StaticFiles(directory="app/static"), name="static")


@app.get("/")
async def index():
    return FileResponse(
        "app/static/index.html",
        headers={"Cache-Control": "no-store"},
    )
