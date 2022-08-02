using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using WpfAppSmetaGraf.ViewModel;

namespace WpfAppSmetaGraf.Model
{
    public class NullValueException : Exception
    {
        public string parName;
        public NullValueException(string s)
        {
            parName = s;
        }
    }
    public class DontHaveExcelException : Exception
    {
        public string parName;
        public DontHaveExcelException(string s)
        {
            parName = s;
        }
    }
    public class ModelWork: ViewModelBase
    {
        private string _userSmeta;
        private string _userOneSmeta;
        private string _userKS;
        private string _userWhereSave;
        private string _userWhereSaveGraph;
        private string _textError = null;
        private int _size;
        private DateTime _dataStart;
        private int _amountDays;
        private int _amountPeople;
        private int _minDays;
        private int _minPeople;
        private int _maxDays;
        private int _maxPeople;
        private List<int> _amountWorker;
        private List<int> _amountWorkDays;
        private bool _exitOrNot;
        public List<int> AmountWorker { get { return _amountWorker; } set { _amountWorker = value; } }
        public List<int> AmountWorkDays { get { return _amountWorkDays; } set { _amountWorkDays = value; } }
        public int FrontSize { get { return _size; } set { _size = value; } }
        public string TextError { get { return _textError; } set { _textError = value; } }
        public string AdressSmeta { get { return _userSmeta; } set { _userSmeta = value; } }
        public string AdressAktKS { get { return _userKS; } set { _userKS = value; } }
        public string AdressSaveSmeta { get { return _userWhereSave; } set { _userWhereSave = value; } }
        public string AddressOneSmeta { get { return _userOneSmeta; } set { _userOneSmeta = value; } }
        public string AdressSaveGraph { get { return _userWhereSaveGraph; } set { _userWhereSaveGraph = value; } }
        public int MinDays { get { return _minDays; } set { _minDays = value; } }
        public int MaxDays { get { return _maxDays; } set { _maxDays = value; } }
        public int MinPeople { get { return _minPeople; } set { _minPeople = value; } }
        public int MaxPeople { get { return _maxPeople; } set { _maxPeople = value; } }
        public bool ExitOrNot { get { return _exitOrNot; } set { _exitOrNot = value; } }

        //выход из EXCEL
        public static void CloseProcess()
        {
            Process[] List;
            List = Process.GetProcessesByName("EXCEL");
            foreach (Process proc in List)
            {
                proc.Kill();
            }
        }
        //закрытие файлов иксель
        public static void CanselProgram()
        {
            if (CheckIt.Instance != null) CheckIt.Instance.Quit();
            CloseProcess();
        }
       //инициализация адресов со сметами, актами и адрес сохранения ведомости
        public void InitializationFormTE(string adressSmeta, string adressAktKS, string adressWhereSave)
        {
            _userSmeta = adressSmeta;
            _userKS = adressAktKS;
            _userWhereSave = adressWhereSave;
        }
        //получение листа числа рабочих для графического приложения
        public List<int> GetAllWorkers()
        {
            _amountWorker = new List<int>();
            int amount = _maxPeople - _minPeople + 1;
            for (int i = 0; i < amount; i++)
            {
                _amountWorker.Add(_minPeople + i);
            }
            return _amountWorker;
        }
        //получение листа числа рабочих дней для графического приложения
        public List<int> GetAllWorkDays()
        {
            _amountWorkDays = new List<int>();
            int amount = _maxDays - _minDays + 1;
            for (int i = 0; i < amount; i++)
            {
                _amountWorkDays.Add(_minDays + i);
            }
            return _amountWorkDays;
        }
        //инициализация адреса сметы и адреса сохранения графика
        public void InitializationFormGraph(string adressOneSmeta, string adressWhereSaveGr)
        {
            _userOneSmeta = adressOneSmeta;
            _userWhereSaveGraph = adressWhereSaveGr;
        }
     
        public void InputDays(int amountDay)
        {
            _amountDays = amountDay;

        }
        public void InputPeople(int amountPeople)
        {
            _amountPeople = amountPeople;

        }
        public void GetInputValueData(DateTime date)
        {
            _dataStart = date;    
        }
        //метод запуска процесса по созданию ведомсти в режиме Эксперт или технадзор
        public void StartProcess(bool flag)
        {
            try
            {
                Worker ob;
                if (flag) ob = new Expert();
                else ob = new Tehnadzor();    
                ob.Initialization(AdressSmeta, AdressAktKS, AdressSaveSmeta);
                ob.ProccessWithDoc(FrontSize, ref _textError);
                       
            }
            catch (DirectoryNotFoundException exc)
            {
                _textError += exc.Message;
            }
            catch (NullValueException exc)
            {
                _textError += exc.parName;
            }
            catch (DontHaveExcelException ex)
            {
                _textError += ex.parName;
            }
            catch (COMException ex)
            {
                _textError += $"{ex.Message} Вы открыли копию сметы, над которой проводится работа программы";
            }
            finally
            {
                CanselProgram();
            }

        }
        //метод запуска процесса по сбору информации для построения графика
        public GraphWork StartChoice()
        {
            GraphWork ob = new GraphWork();
            try
            {
                ob.InitializationGraph(AddressOneSmeta, AdressSaveGraph);
                ob.ProccessGraphFirst();
                _minDays = ob.GetMinDays();
                _maxDays = ob.GetMaxDays();
                _minPeople = ob.GetMinPeople();
                _maxPeople = ob.GetMaxPeople();
            }
            catch (NullValueException exc)
            {
                _textError += exc.parName;
                CanselProgram();
            }
            return ob;
        }
        //производит построение и запись графика
        public void StartGraphRecord(int color, bool flag)
        {
            try
            {
                GraphWork ob = StartChoice();
                ob.ProccessGraph(ref _textError);
                if(flag)
                {
                    _amountPeople = 0;
                    ob.InputDays(ref _amountDays, ref _amountPeople);
                }
                else
                {
                    _amountDays = 0;
                    ob.InputWorkers(_amountPeople, ref _amountDays);
                }
                ob.RecordGraph(_dataStart, _amountDays, _amountPeople, color, ref _textError);
            }
            catch (NullValueException exc)
            {
                _textError += exc.parName;
            }
            finally
            {
                CanselProgram();
            }
        }
     
    }
}
