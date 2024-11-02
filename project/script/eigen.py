import os
import sys
import cv2
import pandas as pd
from database import *
from kivy.app import App
from kivy.uix.screenmanager import Screen
from kivy.lang import Builder
from datetime import datetime

camera = cv2.VideoCapture(0)
detectorFace = cv2.CascadeClassifier("./lib/haarcascade_frontalface_default.xml")
reconhecedor = cv2.face.EigenFaceRecognizer_create()
reconhecedor.read("./lib/classificadorEigen.yml")
largura, altura = 400, 400
font = cv2.FONT_HERSHEY_COMPLEX_SMALL
id_t = ''
conectado, imagem = camera.read()

while True:
    conectado, imagem = camera.read()
    imagem = cv2.flip(imagem, 180)  #flip apenas na imagem colorida
    imagemCinza = cv2.cvtColor(imagem, cv2.COLOR_BGR2GRAY)
    
    # Detecção de faces
    facesDetectadas = detectorFace.detectMultiScale(imagemCinza, minNeighbors=20, minSize=(30, 30), maxSize=(400, 400))

    for (x, y, l, a) in facesDetectadas:
        imagemFace = cv2.resize(imagemCinza[y:y+a, x:x+l], (largura, altura))
        cv2.rectangle(imagem, (x, y), (x + l, y + a), (0, 0, 255), 2)  # Desenha o retângulo

        # Reconhecimento facial
        id, confianca = reconhecedor.predict(imagemFace)
        id_t = id

        cursor.execute("SELECT nome, cargo FROM users WHERE id = ?", (int(id),))
        resultado = cursor.fetchall()

        if not resultado:
            cv2.putText(imagem, "Nenhum user encontrado", (x, y + a + 30), font, 2, (0, 0, 255))
        else:
            nome_user = resultado[0][0]  # Nome
            cargo_user = resultado[0][1]  # Cargo
            cv2.putText(imagem, nome_user, (x, y + a + 30), font, 2, (0, 255, 0))
            
    # Exibição do vídeo com a detecção facial
    cv2.imshow("Reconhecimento Facial", imagem)
    
    # Pressione 'q' para salvar a presença
    
    if cv2.waitKey(1) == ord('q'):
                
        Builder.load_file('../kv_showresults.kv')  # Carrega o arquivo KV do Kivy
        
        class AnswerInput(Screen):
            pass
        class MainApp(App):
            global result

            cursor.execute(
                "SELECT id, nome, cargo, email, cpf FROM users WHERE id = ?", (str(id_t)))
            result = cursor.fetchone()
            print(result)
            nome = "Nome: " + result[1]
            cargo = "Cargo: " + result[2]
            email = "Email: " + result[3]
            cpf = "CPF: " + result[4]
            source_ = './treinamento/pessoa.' + str(result[0]) + '.1.jpg'

                
            def build(self):
                return AnswerInput()
            def on_start(self):
                # Adicionar um temporizador para fechar a tela após alguns segundos
                self.root_window.bind(on_close=self.close_app)
                self.close_after_delay()
            def close_after_delay(self):
                from kivy.clock import Clock
                # Fecha a aplicação após 2 segundos e a reinicia
                Clock.schedule_once(self.close_app, 10)

            def close_app(self, *args):
                App.get_running_app().stop()
                self.restart_app()
                
            def restart_app(self):
                # Reinicia o programa atual
                os.execv(sys.executable, ['python'] + sys.argv)
                
        nome = nome_user if resultado else "Desconhecido"
        cargo = cargo_user if resultado else "Desconhecido"
        timestamp = datetime.now().strftime('%d-%m-%Y %H:%M:%S')

        # Cria o DataFrame com nome, cargo e horário
        df = pd.DataFrame({
            'Nome': [nome],
            'Cargo': [cargo],
            'Data_Hora': [timestamp]
        })

        file_path = 'dash/registro.xlsx'
        
        if os.path.exists(file_path):
            # Adicionar dados ao arquivo existente
            with pd.ExcelWriter(file_path, mode='a', if_sheet_exists='overlay', engine='openpyxl') as writer:
                df.to_excel(writer, index=False, header=False, startrow=writer.sheets['Sheet1'].max_row)
        else:
            # Criar um novo arquivo Excel com os dados
            df.to_excel(file_path, index=False)

        if __name__ == '__main__':
            MainApp().run()
        break

    # Fechar o programa quando a janela for fechada
    if cv2.getWindowProperty("Reconhecimento Facial", cv2.WND_PROP_VISIBLE) < 1:
        break

# Libera a câmera e fecha todas as janelas
camera.release()
cv2.destroyAllWindows()
