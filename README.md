# Sylomer Stock & Availability

Этап 1.1: загрузка Excel-выгрузок из 1С и отображение ключевых показателей в м².

## Что уже реализовано

- Загрузка одного или нескольких Excel-файлов (`.xlsx`, `.xls`) через браузер.
- Дополнительный `select` со списком локальных Excel-файлов из `files-manifest.json` и запуск расчета по выбранным файлам.
- Локальное подключение библиотеки Excel-парсинга (`vendor/xlsx.full.min.js`), без зависимости от внешнего CDN.
- Парсинг первой вкладки файла и автоматический поиск строки заголовков.
- Вывод столбцов:
  - `В наличии м²`
  - `В резерве м²`
  - `Доступно м²`
  - `К обеспечению м²`
  - `Дефицит м²`
  - `Страховой запас м²`
- Пересчет единиц в м²:
  - `пог. м` -> умножение на `1.5`
  - `лист` -> умножение на `1.2`
- Подсчет итогов по всем загруженным строкам.

## Запуск

Достаточно открыть `index.html` в браузере.

Важно: расчет через локальный `select` работает при запуске через HTTP-сервер (например, `python3 -m http.server 8000`), потому что браузерные ограничения блокируют `fetch` из `file://`.

Для локального HTTP-сервера можно использовать:

```bash
python3 -m http.server 8000
```

После этого открыть: <http://localhost:8000>

## Публикация на GitHub и GitHub Pages

В проекте уже есть workflow `.github/workflows/deploy-pages.yml` для автоматического деплоя.

### 1) Первый пуш проекта

Создайте пустой репозиторий на GitHub, затем выполните:

```bash
git init
git add .
git commit -m "Initial version: Excel upload and calculations"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/sylomer-stock-availability.git
git push -u origin main
```

### 2) Включение GitHub Pages

1. Откройте `Settings` -> `Pages` в репозитории.
2. В блоке `Build and deployment` выберите `Source: GitHub Actions`.
3. Дождитесь успешного выполнения workflow `Deploy to GitHub Pages` во вкладке `Actions`.

Сайт будет доступен по адресу:
`https://YOUR_USERNAME.github.io/sylomer-stock-availability/`

### 3) Публикация следующих изменений

```bash
git add .
git commit -m "Update project"
git push
```

После каждого push сайт будет обновляться автоматически.
