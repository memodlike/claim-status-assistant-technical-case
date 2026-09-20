# Claim Status Assistant — технический кейс Product Masters

**Sinoasia B&R / BCC Insurance · КАСКО / страховой случай**  
**Автор:** Karimov Muhamedzhan

После подачи заявления клиенту не всегда понятно, **что уже принято, каких документов не хватает, где находится дело и что делать дальше**.

Я сузил задачу до одного проверяемого сценария и собрал MVP: **статус заявления + чек-лист документов + понятный ответ, который перед отправкой подтверждает сотрудник**.

> ИИ не принимает решение о выплате или отказе. Он помогает структурировать информацию и подготовить понятную коммуникацию.

## Figma

**Прототип:** https://www.figma.com/design/eRYcFzUO9slgs4KrHdICbE

В макете:
- клиентский экран со статусом и документами;
- рабочее место сотрудника;
- экран эксперимента и метрик.

## Готовые файлы

- [PDF — основной отчёт](./artifacts/SABR_Claim_Status_Assistant_Technical_Case_Karimov_Muhamedzhan.pdf)
- [Word — основной отчёт](./artifacts/SABR_Claim_Status_Assistant_Technical_Case_Karimov_Muhamedzhan.docx)
- [Excel — расчётная модель](./artifacts/SABR_Claim_Status_Assistant_Model_Karimov_Muhamedzhan.xlsx)
- [PowerPoint — презентация](./artifacts/SABR_Claim_Status_Assistant_Presentation_Karimov_Muhamedzhan.pptx)
- [ZIP — весь пакет](./artifacts/SABR_Claim_Status_Assistant_Technical_Case_Package.zip)
- [Короткий текст для формы](./artifacts/README_for_Typeform.txt)

## Материалы кейса

- [Проблема, решение и экономика](./case-study.md)
- [Evidence log](./evidence.csv)
- [Расчётная модель в CSV](./model.csv)
- [A/B-план](./experiment.md)
- [Как использовался ИИ](./ai-process.md)
- [Структура презентации](./presentation-outline.md)
- [Описание Figma](./figma.md)
- [Короткое описание решения](./submission-description.txt)

## Что проверяем

**A/B-тест на 2 недели.**

Главная гипотеза: если статус, недостающие документы и следующий шаг видны в одном месте, **повторные обращения по статусу снизятся минимум на 20%** без роста жалоб и ошибок.

## Важное ограничение

Публичные отзывы использованы как **сигнал боли**, а не как статистика всей клиентской базы. Неизвестные внутренние показатели в финансовой модели явно обозначены как допущения и перед реальным запуском должны быть заменены фактическими данными компании.

## Воспроизводимость

Файлы в папке `artifacts/` собираются из [Python-скрипта](./scripts/build_artifacts.py) через GitHub Actions. Это позволяет проверить расчёты и заново собрать пакет без ручных правок.
