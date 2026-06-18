# Rdmo Plugin: Docx Export

This plugin is the follow up to RDMO-Hackerton 2026, group "views".

## Installing

### Installing the plugin (PyPi)

⚠️ No releases yet. You can install the current development version from source. Expect unfinished work there.

### installing the plugin from source

Fresh install
```sh
pip install git+https://github.com/maklem/rdmo-plugins-docx-export.git@main
```

force update from source when version did not change
```sh
pip install git+https://github.com/maklem/rdmo-plugins-docx-export.git@main --force-reinstall --no-deps
```

### Configure RDMO-App to expose plugin

In `config/settings/local.py` add an export plugin. For example
```py
PROJECT_EXPORTS.append(('docx', _('Horizon Europe Docx'), 'rdmo_docx_export.exports.HorizonEuropeDocxExport'))
```

## Development

0. Clone this repository.
0. Setup a venv for development
0. Install this plugin as ediable module into venv.
    ```sh
    pip install --editable path/to/rdmo-plugins-docx-export
    ```
0. Install [RDMO](https://github.com/rdmorganiser/rdmo/) into venv.
0. Configure an [RDMO APP](https://github.com/rdmorganiser/rdmo/), including
    ```py
    DEBUG=True
    # ...
    PROJECT_EXPORTS.append(('docx', _('Horizon Europe Docx'), 'rdmo_docx_export.exports.HorizonEuropeDocxExport'))
    ```
0. Run a local development server
    ```sh
    # Linux
    python3 manage.py runserver

    # Windows with venv
    python.exe manage.py runserver
    ```
    ➡️ Django will (almost) always reload, if a source file changes.
