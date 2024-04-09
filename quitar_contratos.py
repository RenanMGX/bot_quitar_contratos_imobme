from PyQt5 import QtCore, QtGui, QtWidgets
from tkinter.filedialog import askopenfilename
from pandas import read_excel, DataFrame
from getpass import getuser
from time import sleep

from Entities.dados import Dados
from Entities.credenciais import Credential
from Entities.imobme import Imobme

import asyncio
import qasync
import os
import json

class Status:
    def __init__(self) -> None:
        self.__path:str = f"C:/Users/{getuser()}/.bot_quitar_contratos/"
        if not os.path.exists(self.path):
            os.makedirs(self.path)
        
        self.__path_config = self.path + "config.json"
        
        if not os.path.exists(self.path_config):
            with open(self.path_config, 'w', encoding='utf-8')as _file:
                json.dump({"ambiente":"qas"}, _file)
        
    @property
    def path(self) -> str:
        return self.__path
    
    @property
    def path_config(self) -> str:
        return self.__path_config
    
    def read(self) -> dict:
        with open(self.path_config, 'r', encoding="utf-8")as _file:
            return json.load(_file)
        
    def save(self, *, key:str, value:str)-> None:
        param = self.read()
        param[key] = value
        with open(self.path_config, 'w', encoding='utf-8')as _file:
            json.dump(param, _file)


class Ui_MainWindow(object):
    def __init__(self) -> None:
        self.__df:DataFrame = DataFrame()
    
    @property
    def df(self):
        return self.__df
    @df.setter
    def df(self, value:DataFrame):
        if (not isinstance(value, DataFrame)) and (not isinstance(value, Exception)):
            raise TypeError("Apenas Dataframes")
        self.__df = value
    
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.setWindowModality(QtCore.Qt.NonModal) #type: ignore
        MainWindow.setEnabled(True)
        MainWindow.resize(600, 345)
        
        
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(MainWindow.sizePolicy().hasHeightForWidth())
        
        # MainWindow.setSizePolicy(sizePolicy)
        # MainWindow.setMaximumSize(QtCore.QSize(16777215, 16777215))
        # MainWindow.setWindowFilePath("")
        # MainWindow.setToolButtonStyle(QtCore.Qt.ToolButtonIconOnly) #type: ignore
        # MainWindow.setAnimated(True)
        # MainWindow.setDocumentMode(False)
        # MainWindow.setTabShape(QtWidgets.QTabWidget.Rounded)
        # MainWindow.setDockNestingEnabled(False)
        # MainWindow.setDockOptions(QtWidgets.QMainWindow.AllowTabbedDocks|QtWidgets.QMainWindow.AnimatedDocks)
        
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setCursor(QtGui.QCursor(QtCore.Qt.ArrowCursor)) #type: ignore
        self.centralwidget.setAcceptDrops(False)
        self.centralwidget.setLayoutDirection(QtCore.Qt.LeftToRight) #type: ignore
        self.centralwidget.setObjectName("centralwidget")
        
        self.tela = QtWidgets.QStackedWidget(self.centralwidget)
        self.tela.setGeometry(QtCore.QRect(200, 20, 380, 310))
        
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Preferred)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.tela.sizePolicy().hasHeightForWidth())
        
        self.tela.setSizePolicy(sizePolicy)
        self.tela.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.tela.setObjectName("tela")
        
        self.tela_login = QtWidgets.QWidget()
        self.tela_login.setObjectName("Tela_Login")
        
        altura = 200
        
        lateral_barras = 100
        lateral_texto = 50
        lateral_senha = 250
        
        self.descri_login_texto = QtWidgets.QLabel(self.tela_login)
        self.descri_login_texto.setObjectName("Descrição_Login")
        self.descri_login_texto.setWordWrap(True)
        self.descri_login_texto.setAlignment(QtCore.Qt.AlignCenter) #type: ignore
        self.descri_login_texto.setGeometry(QtCore.QRect(lateral_texto-20,altura-150,300,100))
       
        self.text_email = QtWidgets.QLabel(self.tela_login)
        self.text_email.setObjectName("Texto_Email")
        self.text_email.setGeometry(QtCore.QRect(lateral_texto,altura,50,20))
        
        self.barra_email = QtWidgets.QLineEdit(self.tela_login)
        self.barra_email.setObjectName("Email")
        self.barra_email.setGeometry(QtCore.QRect(lateral_barras,altura,230,20))
        self.barra_email.setText(crd.load()['user'])
        
        
        self.text_senha = QtWidgets.QLabel(self.tela_login)
        self.text_senha.setObjectName("Texto_Senha")
        self.text_senha.setGeometry(QtCore.QRect(lateral_texto, altura + 30, 50,20))
        
        self.barra_senha = QtWidgets.QLineEdit(self.tela_login)
        self.barra_senha.setEchoMode(QtWidgets.QLineEdit.Password)
        self.barra_senha.setObjectName("Senha")
        self.barra_senha.setGeometry(QtCore.QRect(lateral_barras, altura + 30, 230, 20))
        self.barra_senha.setText(crd.load()['password'])
        
        
        self.bt_salvar = QtWidgets.QPushButton(self.tela_login)
        self.bt_salvar.setObjectName("Botão_Salvar")
        self.bt_salvar.setGeometry(lateral_senha, altura + 60, 80,25)
        self.bt_salvar.clicked.connect(self.salvar)
        
        self.rd_prd = QtWidgets.QRadioButton(self.tela_login)
        self.rd_prd.setObjectName("CheckBox Produção")
        self.rd_prd.setGeometry(lateral_senha-200, altura + 60, 100,25)
        
        self.rd_qas = QtWidgets.QRadioButton(self.tela_login)
        self.rd_qas.setObjectName("CheckBox Qualidade")
        self.rd_qas.setGeometry(lateral_senha-100, altura + 60, 100,25)
        
        if status.read()['ambiente'] == 'prd':
            self.rd_prd.setChecked(True)
        elif status.read()['ambiente'] == 'qas':
            self.rd_qas.setChecked(True)
        else:
            status.save(key='ambiente', value='qas')
            self.rd_qas.setChecked(True)
        
        self.tela.addWidget(self.tela_login)
                
        self.tela_principal = QtWidgets.QWidget()
        self.tela_principal.setObjectName("Tela_principal")
        
        self.text_principal = QtWidgets.QLabel(self.tela_principal)
        self.text_principal.setObjectName("Texto_area_principal")
        self.text_principal.setGeometry(15,10,350,50)
        self.text_principal.setWordWrap(True)
        self.text_principal.setAlignment(QtCore.Qt.AlignCenter)# type: ignore
        #self.text_principal.setStyleSheet("border: 2px solid black;")
        
        self.bt_carregar_arquivo = QtWidgets.QPushButton(self.tela_principal)
        self.bt_carregar_arquivo.setObjectName("Botão Carregar Arquivo")
        self.bt_carregar_arquivo.setGeometry(140,55,100,25)
        self.bt_carregar_arquivo.clicked.connect(self.carregar_arquivos)
        
        self.retorno_arquivos = QtWidgets.QLabel(self.tela_principal)
        self.retorno_arquivos.setObjectName("Retorno dos arquivos")
        self.retorno_arquivos.setGeometry(15,85,350,150)
        self.retorno_arquivos.setWordWrap(True)
        self.retorno_arquivos.setAlignment(QtCore.Qt.AlignCenter)# type: ignore
        #self.retorno_arquivos.setStyleSheet("border: 2px solid black;")
        
        self.bt_iniciar = QtWidgets.QPushButton(self.tela_principal)
        self.bt_iniciar.setObjectName("Botão Iniciar")
        self.bt_iniciar.setGeometry(115,275,150,25)
        self.bt_iniciar.setVisible(False)
        self.bt_iniciar.clicked.connect(self.iniciar)
        
        self.tela.addWidget(self.tela_principal)
        
        self.verticalLayoutWidget = QtWidgets.QWidget(self.centralwidget)
        self.verticalLayoutWidget.setGeometry(QtCore.QRect(20, 20, 160, 131))
        self.verticalLayoutWidget.setObjectName("verticalLayoutWidget")
        
        self.area_botoes = QtWidgets.QVBoxLayout(self.verticalLayoutWidget)
        self.area_botoes.setSizeConstraint(QtWidgets.QLayout.SetDefaultConstraint) #type: ignore
        self.area_botoes.setContentsMargins(0, 0, 0, 0)
        self.area_botoes.setObjectName("area_botoes")
        
        self.bt_login = QtWidgets.QPushButton(self.verticalLayoutWidget)
        self.bt_login.setObjectName("bt_tela_login")
        self.bt_login.clicked.connect(self.ir_tela_login)
        
        self.area_botoes.addWidget(self.bt_login)
        
        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        self.tela.setCurrentIndex(1)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        self._translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(self._translate("MainWindow", f"{version} - Bot Quitar Contratos"))
        
        self.bt_login.setText(self._translate("MainWindow", "Login"))
        
        self.text_email.setText(self._translate("MainWindow", "Email: "))
        self.text_senha.setText(self._translate("MainWindow", "Senha: "))
        self.bt_salvar.setText(self._translate("MainWindow", "Salvar"))
        
        self.bt_carregar_arquivo.setText(self._translate("MainWindow", "Carregar"))
        
        self.bt_iniciar.setText(self._translate("MainWindow", "Iniciar"))
        
        self.rd_prd.setText(self._translate("MainWindow", "Produção"))
        self.rd_qas.setText(self._translate("MainWindow", "Qualidade"))
        
        self.descri_login_texto.setText(self._translate("MainWindow", "Login\ndigite suas crendenciais do Imobme!"))
        self.text_principal.setText(self._translate("MainWindow", "Carregar Arquivo\nSelecione um arquivo Excel .xlsx"))
    
    #funcionalidades da interface    
    def ir_tela_login(self):
        self.tela.setCurrentIndex(0)
        self.bt_login.disconnect()
        self.bt_login.setText(self._translate("MainWindow", "Voltar"))
        self.bt_login.clicked.connect(self.ir_tela_principal)
        
    def ir_tela_principal(self):
        self.tela.setCurrentIndex(1)
        self.bt_login.disconnect()
        self.bt_login.setText(self._translate("MainWindow", "Login"))
        self.bt_login.clicked.connect(self.ir_tela_login)
    
    #demais script do robo  
    def carregar_arquivos(self):
        self.bt_iniciar.setVisible(False)
        MainWindow.hide()
        try:
            self.retorno_arquivos.setText("")
            try:
                caminho_arquivo:str = askopenfilename()
                if caminho_arquivo == "":
                    raise ValueError("arquivo não selecionado!")
                self.df = Dados.carregar_arquivo(caminho_arquivo)
                self.retorno_arquivos.setText(f"Arquivo Carregado\n{str(caminho_arquivo)}")
                self.bt_iniciar.setVisible(True)
            except Exception as error:
                self.retorno_arquivos.setText(f"{str(type(error))}\n{str(error)}\n{str(caminho_arquivo)}")
                return
        finally:
            MainWindow.show()
            MainWindow.raise_()
            MainWindow.activateWindow()
    
    def salvar(self):
        from PyQt5.QtWidgets import QMessageBox
        
        if self.rd_prd.isChecked():
            status.save(key='ambiente', value='prd')
        elif self.rd_qas.isChecked():
            status.save(key='ambiente', value='qas')
        
           
        crd.save(
            user = self.barra_email.text(),
            password = self.barra_senha.text()
        )
        self.barra_email.setText(crd.load()['user'])
        self.barra_senha.setText(crd.load()['password'])
        
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText("Credenciais salvas com Sucesso!")
        msg.setWindowTitle("Aviso")
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()
    
    def iniciar(self):
        asyncio.create_task(self.start())
        
    async def start(self):
        MainWindow.hide()
        self.bt_iniciar.setVisible(False)
        
        try:
            imobme:Imobme = Imobme(user=crd.load()['user'], password=crd.load()['password'], ambiente=status.read()['ambiente'])
            
            for row,value in self.df.iterrows():
                try:
                    print(imobme.executar(empreendimento=value['Empreendimento'], bloco=value['Bloco'], unidade=value['Unidade']))
                except Exception as error:
                    print(error)
                
                
            self.retorno_arquivos.setText("Concluido!")            
        except Exception as error:
            self.retorno_arquivos.setText(f"{str(type(error))}\n{str(error)}")
        finally:
            MainWindow.show()
            MainWindow.raise_()
            MainWindow.activateWindow()

if __name__ == "__main__":
    version = "v1.0"
    
    status = Status()
    crd = Credential(f"C:/Users/{getuser()}/.bot_quitar_contratos/credencial_imobme_qas.json")
    
    import sys
    app = QtWidgets.QApplication(sys.argv)
    loop = qasync.QEventLoop(app)
    asyncio.set_event_loop(loop)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    loop.run_forever()
