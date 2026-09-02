# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project

**Avaliador de Currículos com IA** — Streamlit web app for CV screening using Google Gemini API. All UI text is in Portuguese (Brazil).

## Running the app

```bash
python -m streamlit run app.py
```

Access at `http://localhost:8501`. If port is busy, Streamlit auto-selects 8502+.

## Installing dependencies

```bash
python -m pip install -r requirements.txt
```

## Architecture

The entire application is a single file (`app.py`) — no modules, no pages. Streamlit reruns the full script on every user interaction.

**Execution order (top to bottom on every rerun):**
1. Constants and JSON persistence helpers (`_ler`, `_gravar`, `carregar_*`, `salvar_*`, `deletar_*`)
2. `st.set_page_config` + session state initialization (score mínimo loaded from `config_app.json`)
3. Pure helper function definitions (`extrair_texto_*`, `extrair_score`, `extrair_veredicto`, `gerar_excel`, `analisar_curriculo`, `comparar_candidatos`, `gerar_roteiro_entrevista`)
4. UI rendering: upload column + saved vagas/perfis sidebar | job description + triagem column
5. Criteria expander with weighted scoring (must sum to 100%)
6. Score mínimo input + Analyze button
7. Batch processing loop (sequential, 4s between CVs, 62s between batches of 11)
8. Results display: dashboard → filters → ranking → per-candidate tabs → comparison

**Session state keys used across reruns:**
- `resultados` — dict keyed by filename, values: `{analise, perguntas, texto_cv}`
- `descricao_vaga`, `score_corte`, `score_corte_usado`
- `ranking_base` — list of `(nome, score, veredicto)` tuples
- `roteiro_{nome}` — on-demand interview script per candidate
- `criterios_sel`, `peso_{criterio}` — criteria widget state

**Gemini API calls:**
- `analisar_curriculo` — 1 call per CV (analysis + interview questions combined, split by `===PERGUNTAS===`)
- `comparar_candidatos` — 1 call for side-by-side comparison
- `gerar_roteiro_entrevista` — 1 call on demand per candidate
- Model: `gemini-3.6-flash`
- Rate limit: batch size 11, 62s pause between batches, 4s between individual CVs

**Persistent JSON files** (excluded from git, created at runtime):
- `vagas_salvas.json` — saved job descriptions `{name: text}`
- `perfis_triagem.json` — saved triagem profiles `{name: text}`
- `criterios_salvos.json` — saved criteria weight configs `{name: {criterio: peso}}`
- `config_app.json` — app config `{score_corte: int}`

**Score enforcement:** Gemini is instructed to apply the minimum score threshold, AND a client-side `re.sub` overrides any "Aprovado" to "Reprovado" for candidates scoring below `score_corte`.

**DOCX extraction** reads both paragraphs AND table cells (CVs often use table layouts).

## Environment

Requires `.env` with:
```
GOOGLE_API_KEY=your_key_here
```

## Git & GitHub

Repository: `https://github.com/fabitarrio/avaliador-curriculos`

**After every code change, commit and push to GitHub:**

```bash
git add app.py
git commit -m "descrição da mudança"
git push origin master
```

If pushing requires authentication, use a Personal Access Token (ghp_...) as the password, or store it in Windows Credential Manager under `git:https://github.com`.

**Files that must never be committed:** `.env`, `vagas_salvas.json`, `perfis_triagem.json`, `criterios_salvos.json`, `config_app.json`, `.streamlit/credentials.toml`

## Regra de Git
Após QUALQUER alteração no projeto, execute automaticamente:
1. git add .
2. git commit -m "descrição da alteração feita"
3. git push origin main


