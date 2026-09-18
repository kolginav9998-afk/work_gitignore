# ПОКАТАК PRIME 2.0.1

Локальный складской учёт (WMS) на LibreOffice Calc + Basic. Одна книга, один каталог товаров,
один внутренний код, один posting engine, один журнал движений, один FIFO, один механизм
возврата, один источник истины — без обязательной зависимости от LibreOffice Base/Firebird.

## Структура репозитория

```
prime_project/
  src/basic/         PRIME_00..14 - исходники Basic-модулей (единственный источник истины для кода)
  src/templates/      Шаблон ODS 1.4.1, на основе которого строится PRIME
  legacy_reference/    Архивные исходники WMS_* версии 1.4.1 - только для сверки, не используются
  tools/              build_ods.py (сборщик), static_check_source.py (статическая проверка исходников)
  tests/              static_checks.py, model_tests.py, functional_smoke.py, fixtures/
  docs/               Этот файл + ARCHITECTURE/DATA_MODEL/POSTING_PROTOCOL/TEST_MATRIX/
                      CHANGELOG/KNOWN_ISSUES/MIGRATION.md, decisions/ (ADR)
  build/              Локальные сборочные артефакты (не коммитятся)
  dist/               Финальные релизные артефакты (не коммитятся, публикуются через GitHub Releases)
```

## Быстрый старт (локальная сборка)

Требуется Python 3 с модулем `uno` (идёт в комплекте с LibreOffice) и установленный
`libreoffice-calc` + `xvfb`:

```bash
sudo apt-get install -y libreoffice-calc xvfb
cd prime_project
python3 tools/build_ods.py --output build/ПОКАТАК_PRIME_2.0.1.ods
python3 tests/static_checks.py build/ПОКАТАК_PRIME_2.0.1.ods
python3 tests/model_tests.py
```

## Документация

- [ARCHITECTURE.md](ARCHITECTURE.md) — архитектурные решения и их обоснование
- [DATA_MODEL.md](DATA_MODEL.md) — схемы всех системных и бизнес-листов
- [POSTING_PROTOCOL.md](POSTING_PROTOCOL.md) — транзакционный протокол проведения документов
- [TEST_MATRIX.md](TEST_MATRIX.md) — что и как проверяется, включая известные ограничения тестов
- [MIGRATION.md](MIGRATION.md) — миграция 1.4.1 → 2.0.0
- [MIGRATION_2.0.0_TO_2.0.1.md](MIGRATION_2.0.0_TO_2.0.1.md) — миграция 2.0.0 → 2.0.1
- [KNOWN_ISSUES.md](KNOWN_ISSUES.md) — известные ограничения текущего релиза
- [CHANGELOG.md](CHANGELOG.md) — история изменений
- [decisions/](decisions/) — architecture decision records (ADR)

## Разработка

См. правила веток/PR/CI в корневом `.github/` репозитория: любые изменения идут через ветки
`claude/*` и pull request в `main`, с обязательными проверками `prime-static`, `prime-build`,
`prime-ods-validate`, `prime-tests`, `prime-libreoffice-linux-smoke`.
