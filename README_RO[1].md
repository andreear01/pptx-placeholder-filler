# PowerPoint Placeholder Filler – tool pentru agentul AI

Acest mini-tool completează automat un PowerPoint care conține placeholders de tip `{{AI_...}}`.

## Ce face

Input:
1. PPTX cu placeholders
2. JSON cu insight-uri generate de agent

Output:
1. PPTX completat
2. `image_prompts.json` pentru placeholders de imagine
3. `fill_report.json` cu verificare:
   - câte placeholders au fost înlocuite
   - ce placeholders au rămas
   - ce keys din JSON nu au fost găsite în PPT

## Structura JSON pe care trebuie să o dea agentul

```json
{
  "{{AI_MX_CATEGORY_MENTIONS_VOLUME}}": "Mentions show a more polarized landscape...",
  "{{AI_MX_OVERALL_INSIGHTS}}": "Line 1.\nLine 2.\nLine 3.",
  "{{AI_IMAGE_MX}}": "Premium smartphone comparison scene, no text, no logos."
}
```

## Rulare locală

```bash
python -m venv .venv
source .venv/bin/activate  # Mac/Linux
# sau
.venv\Scripts\activate     # Windows

pip install -r requirements.txt

uvicorn app:app --reload --port 8000
```

Verificare:
```bash
curl http://localhost:8000/health
```

## Test cu upload de fișier

Deschide în browser:

```text
http://localhost:8000/docs
```

Apoi folosește endpoint-ul:

```text
POST /fill_powerpoint_file
```

Încarci:
- `pptx_file`: PowerPoint-ul cu placeholders
- `replacements_json`: JSON-ul generat de agent

Primești un ZIP cu:
- `completed_report.pptx`
- `image_prompts.json`
- `fill_report.json`

## Securitate

Pentru deploy, setează o cheie API:

```bash
export PPTX_TOOL_API_KEY="your-secret-key"
```

API-ul va cere header:

```text
x-api-key: your-secret-key
```

## Deploy

Poate fi publicat pe:
- Azure Function / Azure Container Apps
- Google Cloud Run
- Render
- server intern aprobat de companie

După deploy, înlocuiești în `openapi_action_schema.yaml`:

```yaml
servers:
  - url: https://YOUR-DOMAIN-HERE.com
```

cu URL-ul real al API-ului.

## Conectare la agent

În GPT/agent builder:
1. mergi la Actions / Tools / Connect external API
2. adaugi schema OpenAPI din `openapi_action_schema.yaml`
3. setezi authentication ca API key, dacă folosești `PPTX_TOOL_API_KEY`
4. testezi endpoint-ul în Preview

## Instrucțiune pentru agent

Adaugă în instrucțiunile agentului:

```text
After generating and validating the placeholder JSON, call the PowerPoint Placeholder Filler tool.
Use the working PPTX file and the validated JSON mapping.
Return the completed PowerPoint file to the user.
If the tool report shows remaining placeholders, report them clearly and do not claim the PPT is final.
```
