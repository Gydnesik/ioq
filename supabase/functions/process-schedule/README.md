# process-schedule

Edge Function принимает от администратора таблицу расписания `.xls`, `.xlsx`, `.ods` или `.csv`.

Пайплайн:
1. браузер отправляет файл в Edge Function только после авторизации администратора;
2. Edge Function читает таблицу через SheetJS (`npm:xlsx`);
3. выбирается наиболее похожий на расписание лист;
4. объединённые ячейки раскрываются в двумерную сетку;
5. сетка конвертируется в чистый CSV;
6. только CSV-текст отправляется в Gemini;
7. нормализованный JSON передаётся в `replace_schedules()`;
8. RPC в одной транзакции удаляет ВСЕ старые строки `schedules` и вставляет только новое расписание.

Gemini получает строгие инструкции игнорировать дежурство, ВПР, примечания и прочую мета-информацию, искать нижнюю основную сетку, читать классы 5А–11Б вертикально, сохранять номер урока и время, сохранять `-` и корректно обрабатывать объединённые ячейки.

## Secrets

В Supabase:

```bash
supabase secrets set GEMINI_API_KEY="YOUR_GEMINI_KEY" GEMINI_MODEL="gemini-3.6-flash"
```

`GEMINI_API_KEY` хранится только на стороне Edge Function.

## Deploy

Перед первым деплоем новой версии обязательно выполни обновлённый `schema.sql` в Supabase SQL Editor — он создаёт RPC `replace_schedules(jsonb)`.

```bash
supabase link --project-ref igbkjkjagkhxpxezjwtj
supabase functions deploy process-schedule --project-ref igbkjkjagkhxpxezjwtj
```
