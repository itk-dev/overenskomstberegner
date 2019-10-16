# Overenskomstberegner

```sh
docker-compose run yarn install
docker-compose run yarn watch
```


## Translations

Run

```sh
DEFAULT_LOCALE=en bin/console translation:update da --force
```

to generate translation files for the `da` locale.

Use [Poedit](Phttps://poedit.net) to edit the generated [`xlf`
files](https://en.wikipedia.org/wiki/XLIFF).
