# edrsr-tool

`edrsr.py` — вибірка судових рішень з відкритих даних ЄДРСР (річні архіви — зі свіжішого з двох джерел: сайт ДСА
https://court.gov.ua/opendata/ — основне, data.gov.ua — якщо його архів новіший на 12+ год чи ДСА не відповідає; `--dzherelo auto|portal|dsa`): рішення у справі, рішення судді за період,
за doc_id, пошук за текстом, коди судів, реєстр Excel і зведений PDF.

Автор: В'ячеслав Плескач. Містить лише код роботи з відкритими даними.

```bash
curl -sSfL https://raw.githubusercontent.com/viacheslavpleskach-cpu/edrsr-tool/main/edrsr.py -o edrsr.py
python3 edrsr.py check
python3 edrsr.py --help
```

Єдина копія скрипта: її використовують скіл «edrsr-court-decisions» і інструменти проєкту «СДІ»
(`sdi-tools/get.sh` завантажує його як `edrsr_vybirka.py`). Зміни вносити лише тут.
