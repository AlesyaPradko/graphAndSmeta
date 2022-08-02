using System;
using WpfAppSmetaGraf.Model;
using System.Windows.Input;
using System.Collections.Generic;
using WpfAppSmetaGraf.Infrastructure;
using Microsoft.Win32;
using System.Windows;
using System.Threading.Tasks;

namespace WpfAppSmetaGraf.ViewModel
{
    internal class MainPageViewModel: ViewModelBase
    {
        public ModelWork _modelWork = new ModelWork();
        public string Adress { get; set; }
        public int IndexFontSize { get; set; }
        public int IndexAmountWorker { get; set; }
        public int IndexAmountDays { get; set; }
        public int IndexColors { get; set; }
        public DateTime DateStart { get; set; }
        private string _textNote= "Добро пожаловать в программу[Помощник эксперта]";
        private bool _testE = false;
        private bool _testT = false;
        private bool _testDay = false;
        private bool _testPeople = false;
        private bool _flag = true;
        private bool _exitOperation = false;
        private bool _daysOrPeople = false;
        private int _colorGet;
        public string TextErrorSm
        {
            get
            { return _modelWork.TextError; }
            set
            {
                _modelWork.TextError = value;
                OnPropertyChanged("TextErrorSm");
            }
        }
        public string TextNote
        {
            get
            { return _textNote; }
            set
            {
                _textNote = value;
                OnPropertyChanged("TextNote");
            }
        }
        public string TextErrorGraph
        {
            get { return _modelWork.TextError; }
            set
            {
                _modelWork.TextError = value;
                OnPropertyChanged("TextErrorGr");
            }
        }
        public List<int> AmountWorkers
        {
            get { return _modelWork.AmountWorker; }
            set
            {
                _modelWork.AmountWorker = value;
                OnPropertyChanged("AmountWorkers");
            }
        }
        public List<int> AmountDays
        {
            get { return _modelWork.AmountWorkDays; }
            set
            {
                _modelWork.AmountWorkDays = value;
                OnPropertyChanged("AmountDays");
            }
        }
        RelayCommand _addAddressFolder;
        public ICommand AddFolder
        {
            get
            {
                if (_addAddressFolder == null)
                    _addAddressFolder = new RelayCommand(ExecuteAddFolderCommand, CanExecuteAddFolderCommand);
                return _addAddressFolder;
            }
        }

        private void ExecuteAddFolderCommand(object parameter)
        {
            System.Windows.Forms.FolderBrowserDialog openFileDlg = new System.Windows.Forms.FolderBrowserDialog();
            var result = openFileDlg.ShowDialog();
            if (result.ToString() != string.Empty)
            {
                Adress = openFileDlg.SelectedPath;
                switch (parameter.ToString())
                {
                    case "ChangeFolderSmeta":
                        _modelWork.AdressSmeta = Adress; break;
                    case "ChangeFolderKS":
                        _modelWork.AdressAktKS = Adress; break;
                    case "SaveCopySmeta":
                        _modelWork.AdressSaveSmeta = Adress; break;
                    case "SaveGraph":
                        _modelWork.AdressSaveGraph = Adress; break;
                }
            }
        }

        private bool CanExecuteAddFolderCommand(object parameter)
        {
            if (parameter.ToString() == "ChangeFolderSmeta" || parameter.ToString() == "ChangeFolderKS"
                || parameter.ToString() == "SaveCopySmeta" || parameter.ToString() == "SaveGraph")
                return true;
            else return false;
        }

        RelayCommand _addAddressSmeta;
        public ICommand AddSmeta
        {
            get
            {
                if (_addAddressSmeta == null)
                    _addAddressSmeta = new RelayCommand(ExecuteAddSmetaCommand, CanExecuteAddSmetaCommand);
                return _addAddressSmeta;
            }
        }
        private void ExecuteAddSmetaCommand(object parameter)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Filter = "Files Excels(*.xlsx;*.csv)|*.xlsx;*.csv";
            if (dlg.ShowDialog() == true)
            {
                _modelWork.AddressOneSmeta = dlg.FileName;
            }
        }
        private bool CanExecuteAddSmetaCommand(object parameter)
        {
            if (parameter.ToString() == "ChangeSmeta")
                return true;
            else return false;
        }
        RelayCommand _addNoteForUser;
        public ICommand AddNote
        {
            get
            {
                if (_addNoteForUser == null)
                    _addNoteForUser = new RelayCommand(ExecuteAddNoteCommand, CanExecuteAddNoteCommand);
                return _addNoteForUser;
            }
        }
        private void ExecuteAddNoteCommand(object parameter)
        {
            switch (parameter.ToString())
            {
                case "NoteSmeta":
                    TextNote = "В режиме [Работы со сметами] пользователь выбирает на своем персональном компьютере папку," +
                        " где располагаются ЛСР, папку, где расположены акты КС-2, папка, куда необходимо сохранить результат работы программы, размер шрифта, " +
                        "форму составляемой ведомости (либо это ведомость в режиме [технадзор], либо это ведомость в режиме [эксперт]), а затем нажимает кнопку «Сформировать ведомость». " +
                        "По указанному пользователем адресу создаются копии ЛСР. Происходит их редактирование, заполнение полученными в процессе работы программы данными, форматирование и сохранение." +
                        " После завершения работы и сохранения ведомостей пользователь информируется о завершении работы, также осуществляется вывод ошибок в файлах ЛСР и актах КС-2 в консоль, при их наличии.";
                    break;
                case "NoteGraph":
                    TextNote = "В режиме [График производства строительных работ] пользователь выбирает на своем персональном компьютере ЛСР в формате xlsx, папку, куда необходимо сохранить результат работы программы. " +
                        "Затем нажимает кнопку «начать работу», после чего обновляются данные в листах комбобоксов «Выберите количество рабочих дней» и «Выберите количество рабочих». " +
                        "Далее пользователь выбирает либо количество человек в бригаде, либо количество рабочих дней, выбирает дату начала работ и цвет графика в исходном документе. " +
                        "После чего пользователь нажимает кнопку «Сформировать график». Далее происходит заполнение исходного файла полученными в процессе работы программы данными," +
                        " происходит заливка графика цветом, форматирование и сохранение. После завершения работы и сохранения ведомостей пользователь информируется о завершении работы, " +
                        "также осуществляется вывод ошибок в файле локального сметного расчета в консоль, при их наличии.";
                    break;
                case "NoteBegin":
                    TextNote = "Добро пожаловать в программу[Помощник эксперта]";
                    break;
            }
        }
        private bool CanExecuteAddNoteCommand(object parameter)
        {
            if (parameter.ToString() == "NoteSmeta"|| parameter.ToString() == "NoteGraph"|| parameter.ToString() == "NoteBegin")
                return true;
            else return false;
        }
        RelayCommand _addSelectMod;
        public ICommand SelectMod
        {
            get
            {
                if (_addSelectMod == null)
                    _addSelectMod = new RelayCommand(ExecuteAddSelectCommand, CanExecuteAddSelectCommand);
                return _addSelectMod;
            }
        }

        private void SelectIndex()
        {
            switch (IndexFontSize)
            {
                case -1:
                    _modelWork.FrontSize = 8; break;
                case 0:
                    _modelWork.FrontSize = 8; break;
                case 1:
                    _modelWork.FrontSize = 9; break;
                case 2:
                    _modelWork.FrontSize = 10; break;
                case 3:
                    _modelWork.FrontSize = 11; break;
                case 4:
                    _modelWork.FrontSize = 12; break;
            }
        }
        private void ExecuteAddSelectCommand(object parameter)
        {
            SelectIndex();
            switch (parameter.ToString())
            {
                case "expert":
                    _testE = true;
                    _testT = false;
                    break;
                case "tehnadzor":
                    _testT = true;
                    _testE=false;
                    break;
                case "workers":
                    if (_daysOrPeople)
                    {
                        _testPeople = true;
                        _testDay = false;
                        AmountWorkers = _modelWork.GetAllWorkers();
                    }
                    else MessageService.ShowExclamation("Нажмите кнопку [Начать работу] чтобы задать интервал рабочих или рабочих дней!");
                    break;
                case "days":
                    if (_daysOrPeople)
                    {
                        _testDay = true;
                        _testPeople = false;
                        AmountDays = _modelWork.GetAllWorkDays();
                    }
                    else MessageService.ShowExclamation("Нажмите кнопку [Начать работу] чтобы задать интервал рабочих или рабочих дней!");
                    break;
            }
        }
        private bool CanExecuteAddSelectCommand(object parameter)
        {
            if (parameter.ToString() == "expert" || parameter.ToString() == "tehnadzor" || parameter.ToString() == "days" || parameter.ToString() == "workers")
                return true;
            else return false;
        }

        private void GetAllErrorSm(string error)
        {
            TextErrorSm = error;
        }
        private void GetAllErrorGr(string error)
        {
            TextErrorGraph  = error;
        }

        RelayCommand _addStartWork;
        public ICommand StartWork
        {
            get
            {
                if (_addStartWork == null)
                    _addStartWork = new RelayCommand(ExecuteAddStartCommand, CanExecuteAddStartCommand);
                return _addStartWork;
            }
        }
        private static int TakeColor(int index)
        {
            int colNum = 0;
            switch (index)
            {
                case 0: colNum = 3; break;
                case 1: colNum = 10; break;
                case 2: colNum = 25; break;
                case 3: colNum = 6; break;
                case 4: colNum = 46; break;
                case 5: colNum = 33; break;
                case 6: colNum = 53; break;
                case 7: colNum = 1; break;
            }
            return colNum;
        }
        private async void WorkWithSmeta(bool test,string note)
        {
            await Task.Factory.StartNew(() =>
            {
                _modelWork.StartProcess(test);
                _flag = true;
                _exitOperation = true;
            });
            if (_exitOperation)
            {
                TakeMassadge(GetAllErrorSm,note);
            }
        }
        private void TakeMassadge(Action<string> act, string note)
        {
            if (_modelWork.TextError.Length == 0)
            {
                MessageService.ShowMessage(note);
            }
            else
            {
                act(_modelWork.TextError);
                MessageService.ShowError("Устраните все ошибки и попробуйте снова");
            }
        }
        private async void ExecuteAddStartCommand(object parameter)
        {
            switch (parameter.ToString())
            {
                case "startSmeta":
                    if (_flag)
                    {
                        GetAllErrorSm(null);
                        GetAllErrorGr(null);
                        if (_modelWork.AdressSmeta != null && _modelWork.AdressSmeta != "" && _modelWork.AdressAktKS != null && _modelWork.AdressAktKS != "" && _modelWork.AdressSaveSmeta != null && _modelWork.AdressSaveSmeta != "")
                        {
                            _flag = false;
                            if (_testE)
                            {
                                WorkWithSmeta(true, "Ведомость эксперта успешно сохранена");
                            }

                            else if (_testT)
                            {
                                WorkWithSmeta(false, "Ведомость технадзора успешно сохранена");
                            }
                            else
                            {
                                MessageService.ShowExclamation("Вы не выбрали режим!");
                                _flag = true;
                            }
                        }
                        else MessageService.ShowExclamation("Вы пытаетесь выполнить работу, но не выбрали папку со сметами, актами КС и для сохранения!");
                    }
                    else MessageService.ShowExclamation("Вы уже запустили процесс обработки. Подождите!");
                    break;
                case "startGraph":
                    if (_flag)
                    {
                        GetAllErrorSm(null);
                        GetAllErrorGr(null);
                        if (_modelWork.AddressOneSmeta != null && _modelWork.AddressOneSmeta != "" && _modelWork.AdressSaveGraph != null && _modelWork.AdressSaveGraph != "" && _daysOrPeople == true)
                        {
                            if (IndexColors == -1) _colorGet = TakeColor(0);
                            else _colorGet = TakeColor(IndexColors);
                            DateTime start = DateStart;
                            _modelWork.GetInputValueData(start);
                            _flag = false;
                            if (_testDay)
                            {
                                await Task.Factory.StartNew(() =>
                                {
                                    if (IndexAmountDays == -1) _modelWork.InputDays(AmountDays[0]);
                                    else _modelWork.InputDays(AmountDays[IndexAmountDays]);
                                    _modelWork.StartGraphRecord(_colorGet,true);
                                    _flag = true;
                                    _exitOperation = true;
                                });
                                if (_exitOperation)
                                {
                                    TakeMassadge(GetAllErrorGr,"График успешно сохранен");
                                }
                            }
                            else if (_testPeople)
                            {

                                await Task.Factory.StartNew(() =>
                                {
                                    if (IndexAmountWorker == -1) _modelWork.InputPeople(AmountWorkers[0]);
                                    else _modelWork.InputPeople(AmountWorkers[IndexAmountWorker]);
                                    _modelWork.StartGraphRecord(_colorGet,false);
                                    _flag = true;
                                    _exitOperation = true;
                                });
                                if (_exitOperation)
                                {
                                    TakeMassadge(GetAllErrorGr, "График успешно сохранен");
                                }
                            }
                            else
                            {
                                MessageService.ShowExclamation("Вы не выбрали количество человек или дней!");
                                _flag = true;
                            }
                        }
                        else MessageService.ShowExclamation("Вы пытаетесь выполнить работу, но не выбрали cмету, и папку для сохранения! Или же вы не нажали кгопку [Начать работу]");

                    }
                    else MessageService.ShowExclamation("Вы уже запустили процесс обработки. Подождите!");
                    break;
            }
        }
        private bool CanExecuteAddStartCommand(object parameter)
        {
            if (parameter.ToString() == "startSmeta" || parameter.ToString() == "startGraph")
                return true;
            else return false;
        }

        RelayCommand _addFirstStartGraph;
        public ICommand AddFirstStart
        {
            get
            {
                if (_addFirstStartGraph == null)
                    _addFirstStartGraph = new RelayCommand(ExecuteAddFirstStartCommand, CanExecuteAddFirstStartCommand);
                return _addFirstStartGraph;
            }
        }
        private async void ExecuteAddFirstStartCommand(object parameter)
        {

            if (_flag)
            {
                if (_modelWork.AddressOneSmeta != null && _modelWork.AdressSaveGraph != null)
                {
                    await Task.Factory.StartNew(()=>
                    {
                        _modelWork.StartChoice();
                        _daysOrPeople = true;
                    });
                } 
                else MessageService.ShowExclamation("Вы пытаетесь выполнить работу, но не выбрали cмету, и папку для сохранения!");
            }
            else MessageService.ShowExclamation("Вы уже запустили процесс обработки. Подождите!");
        }
        private bool CanExecuteAddFirstStartCommand(object parameter)
        {
            if (parameter.ToString() == "FirstStartGraph")
                return true;
            else return false;
        }

        RelayCommand _addExit;
        public ICommand AddExit
        {
            get
            {
                if (_addExit == null)
                    _addExit = new RelayCommand(ExecuteAddExitCommand, CanExecuteAddExitCommand);
                return _addExit;
            }
        }

        private void ExecuteAddExitCommand(object parameter)
        {
            if (parameter.ToString() == "exitSmeta" || parameter.ToString() == "exitGraph")
            {
                if (_flag) 
                { 
                    Application.Current.MainWindow.Close();
                }
                else MessageService.ShowExclamation("Вы не можете выйти, производится работа над файлами");
            }
        }

        private bool CanExecuteAddExitCommand(object parameter)
        {
            if (parameter.ToString() == "exitSmeta" || parameter.ToString() == "exitGraph")
                return true;
            else return false;
        }
    }
}
