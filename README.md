# Leave Application Maker (React SPA)

A client-side React single-page app to edit a leave application, preview it live, and download it as a Microsoft Word `.docx` file.

## Requirements

- Node.js 20+ recommended
- npm

## Run locally

```bash
npm install
npm run dev
```

## Build and preview production output

```bash
npm run build
npm run preview
```

The production-ready static files are generated in `dist/`.

## Deploy to Vercel

1. Push this project to a GitHub/GitLab/Bitbucket repository.
2. Import the repository in Vercel.
3. Framework preset: **Vite**
4. Build command: `npm run build`
5. Output directory: `dist`
6. Deploy.

The included `vercel.json` provides an SPA fallback.

## Deploy to Netlify

- Build command: `npm run build`
- Publish directory: `dist`

For other static hosts, deploy the contents of `dist/`.

## Features

- Editable school, recipient, teacher details, dates, subject, and application text.
- Live letter preview.
- Downloads a real `.docx` generated in the browser.
- Print / Save as PDF.
- Restore original defaults.
- No backend or database required.

## Defaults

The fields are prefilled using the information visible in the supplied handwritten application. Please verify name spellings and dates before submitting.
# Leave-Application-Creator-SPA
