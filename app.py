from flask import Flask, request, render_template, flash, redirect
import pandas as pd
import numpy as np
import os
import re

app = Flask(__name__)


def allowed_file(filename):
    return '.' in filename and \
        filename.rsplit('.', 1)[1].lower() in {'xlsx', 'xlsm', 'csv'}


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        print(request.form['month'])
        month = int(request.form['month'])
        forecast_file = request.files['forecast_file']

        forecast_filename = 'forecast.csv'
        schedule_filename = 'schedule.xlsm'
        forecast_filepath = os.path.join('./', forecast_filename)

        if os.path.isfile(forecast_filepath):
            os.remove(forecast_filepath)

        forecast_file.save(forecast_filepath)

        # Проверяем, был ли предоставлен файл schedule_file
        schedule_file = request.files['schedule_file']
        if schedule_file.filename != '':
            schedule_filepath = os.path.join('./', schedule_filename)

            if os.path.isfile(schedule_filepath):
                os.remove(schedule_filepath)

            schedule_file.save(schedule_filepath)

        schedules = pd.read_excel(schedule_filename, sheet_name='График сменности')
        forecasts = pd.read_csv(forecast_filename, delimiter=';', encoding='Windows-1251').iloc[:, [0, 7]]
        forecasts.iloc[:, 1] = forecasts.iloc[:, 1].apply(lambda x: str(x).replace(',', '.'))

        month_dict = {1: 'январь', 2: 'февраль', 3: 'март', 4: 'апрель', 5: 'май', 6: 'июнь',
                      7: 'июль', 8: 'август', 9: 'сентябрь', 10: 'октябрь', 11: 'ноябрь', 12: 'декабрь'}

        def extract_month(month):
            idx = np.where(schedules.iloc[:, 3].apply(lambda x: month_dict[month] in str(x).lower()))[0][0]
            t_schedule = schedules.iloc[idx + 9:idx + 15, :34]
            t_schedule.iloc[:, 2:] = t_schedule.iloc[:, 2:].map(
                lambda x: 8 if x == 8 else (15 if x == 15 else (8 if x == 12 else 0)))
            return t_schedule

        if month != 1:
            schedule = extract_month(month - 1)
            last_non_empty_column = schedule.columns[(schedule != 0).any()][-1]
            name = schedule[schedule[last_non_empty_column] == 15].iloc[0, 0]
        else:
            schedule = extract_month(month)
            name = schedule[schedule['Unnamed: 3'] == 8].iloc[0, 0]

        schedule = extract_month(month)
        schedule['Unnamed: 2'] = 0
        name_idx = np.where(schedule.iloc[:, 0].apply(lambda x: name.lower() in str(x).lower()))[0][0]
        schedule.iloc[name_idx, 2] = 15

        statistics = {}
        for i in schedule.iloc[:, 0]:
            statistics[i] = []

        pattern = r'^\d{4}\/\d{4}(?: [AC])?$'
        crutch = False
        for i in forecasts.values:
            if not re.match(pattern, str(i[0]).strip()):
                if crutch:
                    break
                else:
                    crutch = True
                    continue
            else:
                crutch = False
            day = i[0].strip().split('/')[0][:2]
            time = i[0].strip().split('/')[0][2:4]
            day = int(day)
            time = int(time)
            if time <= 6:
                name = schedule[(schedule.iloc[:, 2 + day] == 8) & (schedule.iloc[:, 2 + day - 1] == 15)].iloc[0, 0]
                statistics[name].append(float(i[1]))
            else:
                name = schedule[schedule.iloc[:, 2 + day] == 15].iloc[0, 0]
                statistics[name].append(float(i[1]))

        result = {}
        for employee, scores in statistics.items():
            if scores:
                average = sum(scores) / len(scores)
                result[employee] = round(average, 2)
            else:
                result[employee] = 'Нет статистики'
        return render_template('results.html', result=result, month=month_dict[month], year=schedules.iloc[8, 13])
    return render_template('index.html')


if __name__ == '__main__':
    app.run(debug=True)
