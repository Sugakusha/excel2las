#coding utf-8
from PyQt5.Qt import QComboBox, QFormLayout, QLineEdit, QApplication, QPushButton, QWidget, QFileDialog, QMessageBox, QMessageBox, QLabel, QVBoxLayout, QComboBox
from PyQt5.QtCore import Qt
from xlrd import XLRDError
import sys
import pandas as pd
from openpyxl import load_workbook
from os.path import join, split

class ModeError(Exception):
    pass

class Transform:
    
    def __init__(self, data, mode = 'prod'):
        self.data = data.reset_index(drop = True)
        if mode not in ['prod', 'last']:
            raise ModeError("mode must be 'prod' or 'last'.") 
        else:
            self.mode = mode
        
    def __str__(self):
        return 'DataFrame size: {}'.format(self.data.shape)
    
    def fit(self):
        md = sorted(self.data[['md1', 'md2']].values.reshape(-1))#упорядочивание глубин
        sub = []
        for a, b in zip(md[:-1], md[1:]):
            if a == b: continue
            
            #вычисление индексов для которых пересечение 
            #с отрезком из упорядоченных глубин не пусто
            intersect = []
            for i in self.data.index:    
                if a >= self.data.md2[i] or b <= self.data.md1[i]:
                    continue
                else:
                    intersect.append(i)
            
            #пересчет значений модификаторов, как 
            #произведение соотвествующего модификатора 
            #для всех пересекающихся глубин
            props = [] 
            for prop in self.data.columns[3:]:
                val = 1
                if self.mode == 'prod':
                    for i in intersect:
                        val *= self.data[prop][i]
                else:
                    if len(intersect) != 0:
                        val *= self.data[prop][intersect[-1]]
                    
                props.append(val)
            
            sub.append(list(self.data.well.unique()) + [a, b] + props)
            
        return pd.DataFrame(sub, columns = self.data.columns)


class Excel2LasFile:
    
    def __init__(self, path, step, sheet_name = 'Лист1'):
        self.path = path
        self.step = step 
        self.sheet_name = sheet_name

    def __str__(self):
        return 'Excel2LasFile(path = {}, step = {}, sheet_name = {})'.format(self.path, self.step, self.sheet_name)

    def get(self):
        """
        считывае excel файла
        """
        xl = pd.ExcelFile(self.path)
        df = xl.parse(self.sheet_name)
        df.rename(columns = {k: v for k, v in zip(df.columns[:3], ['well', 'md1', 'md2'])}, inplace = True)
        df.well = df.well.astype(str)
        return df
        
    def transform(self, df, mode):
        """
        преобразование DataFrame, дискретизация пересекающихся интервалов
        """
        data = pd.DataFrame(columns = df.columns)
        for well in df.well.unique():
            if len(df[df.well == well]) > 1:
                data = data.append(Transform(df[df.well == well], mode).fit(), ignore_index = True)
            else:
                data = data.append(df[df.well == well], ignore_index = True)
        return data
        
    def post(self, mode = 'prod', path = None): 
        if path is None:
            path = split(self.path)[0] 
        df = self.transform(self.get().fillna(1), mode = mode)
        for well in df.well.unique():
            data = df[df.well == well]
            with open(join(path, '{}.las'.format(well)), 'w') as f:
                    f.write('# MD' + (('  {}') * (len(df.columns[3:]))).format(*df.columns[3:]) + '\n')
                    for i in data.index:
                        start = data.md1[i] 
                        stop = data.md2[i]
                        mods = [data.loc[i, col] for col in data.columns[3:]]
                        while start < stop:
                            f.write((('  {:.5f}') * (len(mods) + 1)).format(start, *mods) + '\n')
                            start += self.step
                        f.write((('  {:.5f}') * (len(mods) + 1) + '\n').format(stop, *mods))


class MessageBox(QMessageBox):
    def __init__(self, title, text, information, critical = True):
        super().__init__()
        if critical:
            self.setIcon(QMessageBox.Critical)
        self.setText(text)
        self.setInformativeText(information)
        self.setWindowTitle(title) 
        self.setStandardButtons(QMessageBox.Ok | QMessageBox.NoButton)
        self.exec_()

        
class MainWindow(QWidget):
    
    def __init__(self, *args, **kwargs):
        super().__init__()
        self.resize(200, 200)
        self.setWindowTitle('Excel2LasFile v0.0.2') 
        self.setWindowFlags(Qt.CustomizeWindowHint | Qt.WindowCloseButtonHint)
        self.Ui()
        self.combobox = 'prod'
        
    def Ui(self):
        self.step_label = QLabel('Step:')
        self._step = QLineEdit()
        self._step.setText('100')
        
        self.sheet_name_label = QLabel('Sheet name:')
        self._sheet_name = QLineEdit()
        self._sheet_name.setText('Лист1')
        
        self.buttonRun = QPushButton()
        self.buttonRun.setText('Run')
        
        self.buttonFile = QPushButton()
        self.buttonFile.setText('Open excel')

        self.buttonSave = QPushButton()
        self.buttonSave.setText('Save las-file')
      
        self.path_label = QLabel('Excel directory')
        self.path_label_save = QLabel('Save directory')
        
        self.cb = QComboBox()
        self.cb.addItems(['prod', 'last'])
        
        
        layout = QFormLayout()
        layout.addWidget(self.step_label)
        layout.addWidget(self._step)
        layout.addWidget(self.sheet_name_label)
        layout.addWidget(self._sheet_name)
        layout.addWidget(self.buttonFile)
        layout.addWidget(self.path_label)
        layout.addWidget(self.buttonSave)
        layout.addWidget(self.path_label_save)
        layout.addWidget(self.buttonRun)
        layout.addWidget(self.cb)
        self.setLayout(layout)
        self.buttonRun.clicked.connect(self.run_button)
        self.buttonFile.clicked.connect(self.openfile)
        self.buttonSave.clicked.connect(self.savefile)
        self.cb.currentIndexChanged.connect(self.selectionchange)
    
    def selectionchange(self):
        self.combobox = self.cb.currentText()

    def backend(self):
        if self._step.text() == '':
            self.step = 100
        else:
            self.step = float(self._step.text())
            assert self.step > 0, 'step must be large then zero'
            
        if self._sheet_name.text() == '':
            self.sheet_name = 'Лист1'
        else:
            self.sheet_name = self._sheet_name.text()     
       
    def openfile(self):
        options = QFileDialog.Options() | QFileDialog.DontUseNativeDialog
        self.path, _ = QFileDialog.getOpenFileName(self, 'openfile', '', 'Excel Files (*.xlsx)', options = options)
        self.path_label.setText(self.path)

    def savefile(self):
        options = QFileDialog.Options() | QFileDialog.DontUseNativeDialog
        self.path_save = QFileDialog.getExistingDirectory(self, options = options)
        if self.path_save == '':
            self.path_save = ''
        self.path_label_save.setText(self.path_save)
        
    def run_button(self):
        try:
            self.selectionchange()
            self.backend()
            obj = Excel2LasFile(path = self.path, step = self.step, sheet_name = self.sheet_name)
            obj.post(path = self.path_save, mode = self.combobox)
        except (FileNotFoundError, AttributeError, IndexError, XLRDError, ValueError) as exc:
            MessageBox(title = 'Error', text = 'Some settings not be found', information = str(exc))
            
            
if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())      
