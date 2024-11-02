### Tools
* Python
* Kivy
* OpenCV

### Hardware
* Raspberry
* Webcam

## This project was made with:

* [Python 3.9.12](https://www.python.org/) 

## how to run the project?

* Create a virtualenv with Python 3.
* Activate virtualenv.
* Install dependencies.
* Create database
* Run migrations.

```
python -m venv .venv
source .venv/Scripts/activate
pip install -r requirements.txt
python project/script/table.py
python project/py_main.py - Executa o reconhecimento facial
python .\dash\dashweb.py - Executa o dashboard
```

## Operação

    1 - Na tela de Cadastro precisa tirar a foto. Clica no Botão "FOTO" e aperte Q para ir tirando as fotos. Precisa apertar umas 10 vezes, pois são 10 fotos.
    2 - Depois preencha os dados do Usuário como nome, cpf, cargo e e-mail. E clica em Salvar.
    3 - Volte para Home e Clica em "Reconhecimento" Abrir reconhecimento e espere aparecer o quadrado vermelho com seu nome e aperte "Q" no teclado.
    4 - Pronto, registro de ponto feito com sucesso.
