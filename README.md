# ប្រព័ន្ធគ្រប់គ្រងរបាយការណ៍ចុះឈ្មោះប្រចាំថ្ងៃ (ខេត្តតាកែវ - PEC21)

គម្រោង Frontend សម្រាប់ Deploy លើ Vercel ដោយតភ្ជាប់ទៅកាន់ Google Apps Script (Google Sheets Backend)។

## រចនាសម្ព័ន្ធឯកសារ
- `index.html` : ឯកសារទំព័រដើម Frontend Web App (អក្សរតូច ដើម្បីជៀសវាង 404 លើ Vercel)
- `gas-api-wrapper.js` : API Proxy Polyfill បម្លែង `google.script.run` ទៅជា `fetch()` ស្វ័យប្រវត្តិ
- `vercel.json` : កំណត់ Routing និង Rewrites លើ Vercel
- `Code.gs` : កូដ `doPost(e)` សម្រាប់ដាក់ក្នុង Google Apps Script
