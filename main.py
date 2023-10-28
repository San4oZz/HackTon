import requests
import json
import math
import aspose.slides as slides
import aspose.pydrawing as drawing


query = """query {
  books {
    edges {
      node {
        id
        displayName
        year
        authors {
          gender
        }
      }
    }
  }
}
"""
url = 'https://api.siberiana.online/graphql'
request = requests.post(url, json={'query': query})
json_data = json.loads(request.text)
df_data = json_data['data']['books']['edges']
all_books_count = 0
dict_year = {'19': 0, '20': 0, '21':0}
dict_gender = {'female': 0, "male": 0, "anon": 0}

for i in df_data:
    year = i['node']['year']
    dict_year[str(year + 100)[:2]] += 1

for i in df_data:
    gender = i['node']['authors']
    if len(gender) != 0:
        dict_gender[gender[0]['gender']] += 1
    else:
        dict_gender['anon'] += 1
    all_books_count += 1

percent_gender = [math.ceil(dict_gender['female'] / all_books_count * 100),
                  math.ceil(dict_gender['male'] / all_books_count * 100),
                  math.ceil(dict_gender['anon'] / all_books_count * 100)]
percent_year = [math.ceil(dict_year['19'] / all_books_count * 100), math.ceil(dict_year['20'] / all_books_count * 100),
                math.ceil(dict_year['21'] / all_books_count * 100)]
count_key = 0

for i, j in dict_gender.items():
    print(i, "\t", j * "=", "\t", percent_gender[count_key], "%")
    count_key += 1

count_key = 0
print('\n')

for i, j in dict_year.items():
    print(i, "век \t", j * "=", "\t", percent_year[count_key], "%")
    count_key += 1


with slides.Presentation() as pres:

    sld = pres.slides[0]

    chart = sld.shapes.add_chart(slides.charts.ChartType.PIE, 100, 100, 400, 400)

    chart.chart_title.add_text_frame_for_overriding("Пол автора")

    chart.chart_title.height = 20
    chart.has_title = True

    chart.chart_data.series[0].labels.default_data_label_format.show_value = True

    defaultWorksheetIndex = 0

    fact = chart.chart_data.chart_data_workbook

    chart.chart_data.series.clear()
    chart.chart_data.categories.clear()

    chart.chart_data.categories.add(fact.get_cell(0, 1, 0, "Женщина"))
    chart.chart_data.categories.add(fact.get_cell(0, 2, 0, "Мужчина"))
    chart.chart_data.categories.add(fact.get_cell(0, 3, 0, "Не указано"))

    series = chart.chart_data.series.add(fact.get_cell(0, 0, 0, 0), chart.type)

    series.data_points.add_data_point_for_pie_series(fact.get_cell(0, 1, 1, percent_gender[0]))
    series.data_points.add_data_point_for_pie_series(fact.get_cell(0, 2, 1, percent_gender[1]))
    series.data_points.add_data_point_for_pie_series(fact.get_cell(0, 3, 1, percent_gender[2]))

    chart.chart_data.series_groups[0].is_color_varied = True

    point = series.data_points[0]
    point.format.fill.fill_type = slides.FillType.SOLID
    point.format.fill.solid_fill_color.color = drawing.Color.orange

    point.format.line.fill_format.fill_type = slides.FillType.SOLID
    point.format.line.fill_format.solid_fill_color.color = drawing.Color.gray
    point.format.line.width = 2.0

    point1 = series.data_points[1]
    point1.format.fill.fill_type = slides.FillType.SOLID
    point1.format.fill.solid_fill_color.color = drawing.Color.blue_violet

    point1.format.line.fill_format.fill_type = slides.FillType.SOLID
    point1.format.line.fill_format.solid_fill_color.color = drawing.Color.blue
    point1.format.line.width = 2.0

    point2 = series.data_points[2]
    point2.format.fill.fill_type = slides.FillType.SOLID
    point2.format.fill.solid_fill_color.color = drawing.Color.yellow_green

    point2.format.line.fill_format.fill_type = slides.FillType.SOLID
    point2.format.line.fill_format.solid_fill_color.color = drawing.Color.red
    point2.format.line.width = 2.0

    lbl1 = series.data_points[0].label

    lbl1.data_label_format.show_value = True

    lbl2 = series.data_points[1].label
    lbl2.data_label_format.show_value = True

    lbl3 = series.data_points[2].label
    lbl3.data_label_format.show_value = True

    chart.chart_data.series_groups[0].first_slice_angle = 180
    pres.save("create-presentation.pptx", slides.export.SaveFormat.PPTX)

with slides.Presentation() as pres:

    sld = pres.slides[0]

    chart = sld.shapes.add_chart(slides.charts.ChartType.PIE_3D, 100, 100, 400, 400)

    chart.chart_title.add_text_frame_for_overriding("Век написания книги")

    chart.chart_title.height = 20
    chart.has_title = True

    chart.chart_data.series[0].labels.default_data_label_format.show_value = True

    defaultWorksheetIndex = 0

    fact = chart.chart_data.chart_data_workbook

    chart.chart_data.series.clear()
    chart.chart_data.categories.clear()

    chart.chart_data.categories.add(fact.get_cell(0, 1, 0, "19 век"))
    chart.chart_data.categories.add(fact.get_cell(0, 2, 0, "20 век"))
    chart.chart_data.categories.add(fact.get_cell(0, 3, 0, "21 век"))

    series = chart.chart_data.series.add(fact.get_cell(0, 0, 0, 0), chart.type)

    series.data_points.add_data_point_for_pie_series(fact.get_cell(0, 1, 1, percent_year[0]))
    series.data_points.add_data_point_for_pie_series(fact.get_cell(0, 2, 1, percent_year[1]))
    series.data_points.add_data_point_for_pie_series(fact.get_cell(0, 3, 1, percent_year[2]))

    chart.chart_data.series_groups[0].is_color_varied = True

    point = series.data_points[0]
    point.format.fill.fill_type = slides.FillType.SOLID
    point.format.fill.solid_fill_color.color = drawing.Color.orange

    point.format.line.fill_format.fill_type = slides.FillType.SOLID
    point.format.line.fill_format.solid_fill_color.color = drawing.Color.gray
    point.format.line.width = 2.0

    point1 = series.data_points[1]
    point1.format.fill.fill_type = slides.FillType.SOLID
    point1.format.fill.solid_fill_color.color = drawing.Color.blue_violet

    point1.format.line.fill_format.fill_type = slides.FillType.SOLID
    point1.format.line.fill_format.solid_fill_color.color = drawing.Color.blue
    point1.format.line.width = 2.0

    point2 = series.data_points[2]
    point2.format.fill.fill_type = slides.FillType.SOLID
    point2.format.fill.solid_fill_color.color = drawing.Color.yellow_green

    point2.format.line.fill_format.fill_type = slides.FillType.SOLID
    point2.format.line.fill_format.solid_fill_color.color = drawing.Color.red
    point2.format.line.width = 2.0

    lbl1 = series.data_points[0].label

    lbl1.data_label_format.show_value = True

    lbl2 = series.data_points[1].label
    lbl2.data_label_format.show_value = True

    lbl3 = series.data_points[2].label
    lbl3.data_label_format.show_value = True

    chart.chart_data.series_groups[0].first_slice_angle = 180
    pres.save("create-presentation1.pptx", slides.export.SaveFormat.PPTX)
