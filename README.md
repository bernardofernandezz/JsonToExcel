# Conversor JSON para Excel

Uma aplicação FastAPI que converte arquivos JSON para Excel com formatação profissional.

## 🚀 Passo a Passo para Execução

### 1. Clone o repositório

- git clone https://github.com/seu-usuario/JsonToExcel.git
- cd JsonToExcel

### 2. Crie um ambiente virtual

# Windows
python -m venv venv
venv\Scripts\activate

# Linux/Mac
python3 -m venv venv
source venv/bin/activate

### 3. Instale as dependências

pip install fastapi uvicorn pandas openpyxl jinja2 python-multipart

### 4. Execute o servidor (escolha uma das opções)

# Opção 1 - Usando python
python main.py

# Opção 2 - Usando uvicorn
uvicorn main:app --reload

### 5. Acesse a aplicação
- Interface web: http://localhost:8000
- Documentação da API: http://localhost:8000/docs

## 📝 Como usar

### Via Interface Web
1. Acesse http://localhost:8000
2. Você pode:
   - Fazer upload de um arquivo JSON
   - Ou colar diretamente o conteúdo JSON no campo de texto
3. Clique em "Converter" e o arquivo Excel será baixado automaticamente

### Via API
A API oferece dois endpoints:

1. Converter arquivo JSON:
```bash
curl -X POST "http://localhost:8000/convert/file/" \
     -H "accept: application/json" \
     -H "Content-Type: multipart/form-data" \
     -F "json_file=@seu_arquivo.json"
```

2. Converter texto JSON:
```bash
curl -X POST "http://localhost:8000/convert/json/" \
     -H "accept: application/json" \
     -H "Content-Type: application/x-www-form-urlencoded" \
     -d "json_text={\"nome\":\"exemplo\",\"idade\":30}"
```

## ✨ Funcionalidades
- Conversão de arquivo JSON para Excel
- Conversão de texto JSON para Excel
- Formatação profissional do Excel
- Suporte a JSON aninhado
- Mapeamento automático de colunas
- Interface web amigável

## 🛠️ Tecnologias Utilizadas
- FastAPI
- Pandas
- OpenPyXL
- Jinja2
- Python-Multipart

## ⚠️ Observações Importantes
- Certifique-se de que a porta 8000 não está sendo usada
- A pasta `uploads` será criada automaticamente
- Mantenha o arquivo `.gitignore` para excluir arquivos desnecessários
- Python 3.9 ou superior é recomendado

## 📄 Licença
Este projeto está sob a licença MIT.
