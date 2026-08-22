# LeverEdge Excel Converter

A browser-only React + Vite tool that processes LeverEdge sales data (Sales Register, Bill Wise Sales, Sale Return) and converts it into formatted Excel files for accounting import.

## Stack
- **React 19** + **TypeScript**
- **Vite 8** (dev server on port 5000)
- **TanStack Router** (file-based routing under `src/routes/`)
- **Tailwind CSS v4** + **shadcn/ui** (Radix primitives)
- **xlsx** — client-side Excel parsing and generation
- **Supabase** — optional backend (env vars in `.env`)

## Running the app
```bash
npm install
npm run dev        # starts on http://localhost:5000
```

The configured Replit workflow (`Start application`) runs `npm run dev` automatically.

## Routes
| File | URL | Description |
|------|-----|-------------|
| `src/routes/index.tsx` | `/` | Excel Converter (main page) |
| `src/routes/matching.tsx` | `/matching` | Data Matching |
| `src/routes/view.tsx` | `/view` | Data View |

## Key notes
- 100% browser-based — no server, no upload, no storage. Files never leave the user's machine.
- Item Master Database persists in `localStorage`.
- All Excel parsing/generation is done client-side with the `xlsx` library.
- The app is connected to [Lovable](https://lovable.dev) — avoid force-pushing or rebasing published commits.

## User preferences
