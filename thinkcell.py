from thinkcellbuilder import Presentation, Template
import pandas as pd
from datetime import datetime

slide = Template("APR Month End_Digital Performance Update - Copy_Factspan_May (2).pptx")

class Thinkcell:
    def update_chart(self, chart_name, df, output_file_name):
        slide.add_chart_from_dataframe(
        name=chart_name,
        dataframe=df,
        )
        presentation = Presentation()
        presentation.add_template(slide)
        output_file = output_file_name
        try:
            presentation.save_ppttc(output_file)
        except AttributeError:
            try:
                presentation.export(output_file)
            except AttributeError:
                try:
                    presentation.write(output_file)
                except AttributeError as e:
                    print(f"Failed to save presentation: {e}")

        print(f"Presentation saved as {output_file}")

