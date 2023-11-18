import click

from pptx_slide_generator import main, profile


@click.group()
def cli():
    pass


@click.command()
@click.option('--excel-path', type=click.Path(exists=True), required=True, help='Path to the Excel file')
@click.option('--pptx-path', type=click.Path(exists=True), required=True, help='Path to the PowerPoint file')
@click.option('--svg-path', type=click.Path(exists=True), required=True, help='Path to the SVG file')
@click.option('--result-path', type=click.Path(exists=False), required=True, help='Path to the result file')
@click.option('--excel-section-value', type=str, required=True, help='Value for the filtered section in the Excel file')
@click.option('--excel-column-room', type=str, default="Raumnr", help='Column name for the room in the Excel file')
@click.option('--excel-column-layout', type=str, default="Layout", help='Column name for the layout in the Excel file')
@click.option('--excel-column-section', type=str, default="Bereich",
              help='Column name for the section in the Excel file')
@click.option('--excel-column-relevant', type=str, default="Schild",
              help='Column name for the relevant rows in the Excel file')
@click.option('--pptx-slide-idx', type=int, default=0,
              help='Index of the slide in the PowerPoint file containing the SVG')
@click.option('--pptx-shape-exclude', type=str, default="IGNORE", help='Shape exclusion pattern in the PowerPoint file')
@click.option('--pptx-shape-prefix', type=str, default="Raum_", help='Shape prefix pattern in the PowerPoint file')
@click.option('--svg-name-attribute', type=str, default="id", help='Name attribute for the SVG file')
def run(**kwargs):
    """Main entry point to execute generate slides from svg and excel input."""
    main.run(**kwargs)
    profile.add_previous_profile(kwargs)


@click.command()
@click.option('--profile-name', type=str, required=True)
@click.option('--excel-path', type=str, required=True, help='Path to the Excel file')
@click.option('--pptx-path', type=str, required=True, help='Path to the PowerPoint file')
@click.option('--svg-path', type=str, required=True, help='Path to the SVG file')
@click.option('--result-path', type=str, required=True, help='Path to the result file')
@click.option('--excel-section-value', type=str, required=True, help='Value for the filtered section in the Excel file')
@click.option('--excel-column-room', type=str, default="Raumnr", help='Column name for the room in the Excel file')
@click.option('--excel-column-layout', type=str, default="Layout", help='Column name for the layout in the Excel file')
@click.option('--excel-column-section', type=str, default="Bereich",
              help='Column name for the section in the Excel file')
@click.option('--excel-column-relevant', type=str, default="Schild",
              help='Column name for the relevant rows in the Excel file')
@click.option('--pptx-slide-idx', type=int, default=0,
              help='Index of the slide in the PowerPoint file containing the SVG')
@click.option('--pptx-shape-exclude', type=str, default="IGNORE", help='Shape exclusion pattern in the PowerPoint file')
@click.option('--pptx-shape-prefix', type=str, default="Raum_", help='Shape prefix pattern in the PowerPoint file')
@click.option('--svg-name-attribute', type=str, default="id", help='Name attribute for the SVG file')
def create_profile(profile_name: str, **kwargs):
    """Create a profile and make it reusable for later usage."""
    profile.add_profile(profile_name, kwargs)


@click.command()
def create_default_profile():
    """Helper to create a first example profile.

    """

    values = dict(excel_path="PathToExcelFile",
                  pptx_path="PathToPptxFile",
                  svg_path="PathToSvgFile",
                  result_path="PathToResultPptxFile")

    profile.add_profile("default", values)


@click.command()
@click.option('--profile-name', type=str, required=True)
@click.option('--excel-path', type=str, required=False, help='Path to the Excel file')
@click.option('--pptx-path', type=str, required=False, help='Path to the PowerPoint file')
@click.option('--svg-path', type=str, required=False, help='Path to the SVG file')
@click.option('--result-path', type=str, required=False, help='Path to the result file')
@click.option('--excel-section-value', type=str, required=False,
              help='Value for the filtered section in the Excel file')
@click.option('--excel-column-room', type=str, default="Raumnr", help='Column name for the room in the Excel file')
@click.option('--excel-column-layout', type=str, default="Layout", help='Column name for the layout in the Excel file')
@click.option('--excel-column-section', type=str, default="Bereich",
              help='Column name for the section in the Excel file')
@click.option('--excel-column-relevant', type=str, default="Schild",
              help='Column name for the relevant rows in the Excel file')
@click.option('--pptx-slide-idx', type=int, default=0,
              help='Index of the slide in the PowerPoint file containing the SVG')
@click.option('--pptx-shape-exclude', type=str, default="IGNORE", help='Shape exclusion pattern in the PowerPoint file')
@click.option('--pptx-shape-prefix', type=str, default="Raum_", help='Shape prefix pattern in the PowerPoint file')
@click.option('--svg-name-attribute', type=str, default="id", help='Name attribute for the SVG file')
def run_profile(profile_name: str, **kwargs):
    """Run a parameter configuration specified by given profile.

    """

    values = profile.read_profile(profile_name)
    kwargs.update(values)
    main.run(**kwargs)


@click.command()
def run_previous_configuration():
    """Rerun previous parameter configurations.

    """

    values = profile.read_previous_profile()
    main.run(**values)


cli.add_command(run)
cli.add_command(create_profile)
cli.add_command(run_profile)
cli.add_command(run_previous_configuration)
cli.add_command(create_default_profile)

if __name__ == '__main__':
    cli()
