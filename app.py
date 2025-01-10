from flask import Flask, render_template, request, send_file
import os
import pandas as pd
import numpy as np
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # максимальный размер файла 16MB

# Создаем папку для загрузки файлов, если её нет
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'files[]' not in request.files:
            return 'Нет загруженных файлов', 400
        
        files = request.files.getlist('files[]')
        search_value = request.form['search_value']
        sheet_name = request.form['sheet_name']
        price_column = request.form['price_column']
        
        # Очищаем папку uploads
        for f in os.listdir(app.config['UPLOAD_FOLDER']):
            os.remove(os.path.join(app.config['UPLOAD_FOLDER'], f))
        
        # Сохраняем загруженные файлы
        excel_files = []
        data_new = []
        for file in files:
            if file.filename.endswith('.xlsx'):
                filename = secure_filename(file.filename)
                file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
                excel_files.append(filename)
                # Получаем дату из имени файла
                data = filename[filename.find('(') + 1: filename.rfind(')')]
                data_new.append(data)
        
        # Создаем пустую итоговую таблицу
        df_res = pd.DataFrame()
        
        results = []
        for i in range(len(excel_files)):
            filename = os.path.join(app.config['UPLOAD_FOLDER'], excel_files[i])
            df = pd.read_excel(filename, sheet_name=sheet_name)
            idx = np.where(df == search_value)
            if len(idx[0]) > 0:  # Проверяем, найдено ли значение
                value = df.loc[idx[0], price_column]
                value_new = value.tolist()
                results.append({
                    'date': data_new[i],
                    'value': search_value,
                    'price': value_new[0]
                })
                new_row = [data_new[i], search_value, value_new[0]]
                new_df = pd.DataFrame([new_row], columns=['Дата', 'Наименование', 'Цена'])
                df_res = pd.concat([df_res, new_df], axis=0, ignore_index=True)
        
        # Сохраняем результат
        output_file = os.path.join(app.config['UPLOAD_FOLDER'], 'output.xlsx')
        df_res.to_excel(output_file, index=False)
        
        return render_template('result.html', results=results, output_file='output.xlsx')
    
    return render_template('index.html')

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(os.path.join(app.config['UPLOAD_FOLDER'], filename),
                    as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True) 