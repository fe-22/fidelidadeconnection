from flask import Flask, request, send_from_directory, render_template, redirect, url_for
import pandas as pd
import os

app = Flask(__name__)

# Página principal
@app.route("/")
def index():
    try:
        return render_template('cad_mem.html')
    except Exception as e:
        return f"Erro ao carregar a página principal: {e}"

@app.route('/favicon.ico')
def favicon():
    return send_from_directory(os.path.join(app.root_path,'favicon.ico'), mimetype='image/vnd.microsoft.icon')


# Diretório onde os arquivos Excel serão salvos
UPLOAD_FOLDER = os.path.join('static', 'uploads')
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Garante que o diretório de uploads existe
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Upload de arquivo
@app.route('/uploads', methods=['POST'])
def upload_file():
    try:
        if 'file' not in request.files:
            return 'No file part'
        file = request.files['file']
        if file.filename == '':
            return 'No selected file'
        if not file.filename.endswith('.xlsx'):
            return 'Invalid file type'
        filename = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(filename)
        return 'File uploaded successfully'
    except Exception as e:
        return str(e)

# Geração do arquivo Excel
@app.route('/generate_excel', methods=['POST'])
def generate_excel():
    try:
        data = request.form.to_dict()

        # Processar os dados para gerar o DataFrame
        df = pd.DataFrame([data])

        # Caminho completo para o arquivo Excel
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], 'dados_formulario.xlsx')

        # Verificar se o arquivo já existe
        if os.path.exists(excel_path):
            # Carregar o arquivo existente
            df_existing = pd.read_excel(excel_path)
            # Concatenar o DataFrame existente com o novo DataFrame
            df_combined = pd.concat([df_existing, df], ignore_index=True)
            # Salvar os dados no arquivo Excel
            df_combined.to_excel(excel_path, index=False)
        else:
            # Se o arquivo não existir, criar um novo com os dados
            df.to_excel(excel_path, index=False)

        return 'Excel gerado com sucesso!'
    except Exception as e:
        return str(e)

@app.route('/uploads_excel')
def download_excel():
    try:
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], 'dados_formulario.xlsx')
        return send_from_directory(excel_path, as_attachment=True, download_name='dados_formulario.xlsx')
    except Exception as e:
        return str(e)

@app.route('/result', methods=['GET'])
def result():
    try:
        # Carregar os dados da planilha
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], 'dados_formulario.xlsx')
        if os.path.exists(excel_path):
            df = pd.read_excel(excel_path)
            # Obter os últimos 10 cadastros
            ultimos_10_cadastros = df.tail(10)
            return render_template('pes.html', cadastros=ultimos_10_cadastros.to_dict(orient='records'))
        else:
            return redirect(url_for('index'))
    except Exception as e:
        return str(e)

@app.route('/404')
def render_404_page():
    try:
        return render_template('404.html')
    except Exception as e:
        return str(e)

if __name__ == "__main__":
    app.run(debug=True)
