# Door Plate Generator

This package contains a narrowly focused command-line interface (CLI) to generate door plates from door plans (SVG), master layouts (PPTX) and master data (XLSX).

## Features

- **Generate Door Plates**: Automatically create powerpoint slides representing door plates.
- **Profile Management**: Create, run, and reuse profiles for different configurations.
- **Command-Line Interface**: Easy-to-use commands for generating slides and managing profiles.

## Installation

You require a [git](https://git-scm.com/), [python interpreter](https://www.python.org/downloads/) (>=3.9) and [poetry](https://python-poetry.org/). To install the application, run the following:

```bash
git clone https://github.com/mansenfranzen/door_plate_generator.git
cd door_plate_generator
poetry install
```

By default, [`poetry.lock`](https://python-poetry.org/docs/basic-usage/#installing-with-poetrylock) is used to provide pinned dependencies for reproducible environment.

## Usage

To activate the python environment, use `poetry shell` first to enable the CLI entry points.

### Run

The package provides a main entry point via `door-plate-generator run` which accepts many different parameters:

```
  --excel-path PATH             Path to the Master Excel File  [required]
  --pptx-path PATH              Path to the PowerPoint Master Layout file [required]
  --svg-path PATH               Path to the SVG Door plans file  [required]
  --result-path PATH            Path to the result file  [required]
  --excel-section-value TEXT    Value for the filtered section in the Master Excel File  [required]
  --excel-column-room TEXT      Column name for the room in the Master Excel File
  --excel-column-layout TEXT    Column name for the layout in the Master Excel File
  --excel-column-section TEXT   Column name for the section in the Master Excel File
  --excel-column-relevant TEXT  Column name for the relevant rows in the Master Excel File
  --pptx-slide-idx INTEGER      Index of the slide in the PowerPoint Master Layout file containing the SVG
  --pptx-shape-exclude TEXT     Shape exclusion pattern in the PowerPoint Master Layout file
  --pptx-shape-prefix TEXT      Shape prefix pattern in the PowerPoint Master Layout file
  --svg-name-attribute TEXT     Name attribute for the SVG Door plans file
  --help                        Show this message and exit.
```

#### Example

```bash
door-plate-generator run \
  --excel-path "path/to/excel.xlsx" \
  --pptx-path "path/to/template.pptx" \
  --svg-path "path/to/image.svg" \
  --result-path "path/to/result.pptx" 
```

### Rerun

To rerun the last configuration, use `door-plate-generator rerun`.

### Profiles

Supplying all parameters for each invocation is tedious. Hence, you can specify reusable profiles with following subcommands for `door_plate_generator profile`: 

- `create`: Create and save a profile with specific configurations.
- `instantiate`: Helper to create a first example profile.
- `execute`: Helper to create a first example profile name.

### Help

Each subcommand contains a help page which can be accessed via `--help` like:

```bash
door-plate-generator profile create --help
```

## Notes

- **Logging**: Warnings regarding inconsistencies between SVG, XLSX and PPTX are written into a `logs` folder in the current working directory. The folder will be created if it doesn't exist. Moreover, infos and warnings are also shown in stdout.
- **Profiles**: Profiles are stored in `%HOME%/door_plate_generator.json` and can be edited there directly.

## License

This project is licensed under the [MIT License](https://en.wikipedia.org/wiki/MIT_License).
