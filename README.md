# SWP Assistant

AI-powered drafting and revision tool for Standard Work Procedures (SWPs). Built as a prototype to help clinical staff turn rough operational notes into formatted SWPs, and to revise existing procedures faster.

## What it does

A clinician describes a process in plain language. The tool drafts a formatted SWP with sections, numbered steps, role assignments, and exception handling. The clinician edits in the browser; the tool revises against the edits and produces a final document.

The model runs against OpenAI's GPT API. A serverless proxy on Vercel holds the API key; the browser never sees it.

## Stack

React + Vite frontend. Vercel serverless function (`api/chat`) as the GPT proxy. ESLint for code style.

## Run locally

```bash
git clone https://github.com/bwcaputo/npm-create-vite-latest.git
cd npm-create-vite-latest
npm install

# Create .env.local (gitignored) with your OpenAI API key:
# OPENAI_API_KEY=sk-...

npm run dev
```

The Vite dev server runs at http://localhost:5173. The serverless proxy is invoked at `/api/chat`.

## Deployment

The production app is deployed at [npm-create-vite-latest.vercel.app](https://npm-create-vite-latest.vercel.app). Push to `main` triggers a Vercel deploy. Set `OPENAI_API_KEY` as a Vercel project environment variable.

## Origin

Built as part of a DU MBA social good project working with a metropolitan health system on clinical workflow automation.

## License

[MIT](LICENSE)
