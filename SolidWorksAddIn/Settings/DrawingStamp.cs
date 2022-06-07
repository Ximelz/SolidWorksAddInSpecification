using System;

namespace Solidworks_PDM_Specification
{
    /// <summary>
    /// Класс хранящий данные о штампе основной надписи.
    /// </summary>
    public class DrawingStamp : Element
    {
        public DrawingStamp()
        {
            developer = "";
            checker = "";
            normativeControl = "";
            approver = "";
            Litera = "";
            InvNumbDupl = "";
            InvNumbOrigin = "";
            ReplacedInvNumb = "";
            ReferenceNumb = "";
            PrimaryApplication = "";
            FilePath = "";
            usePdmFlag = false;
            Configuration = "";
        }

        private string developer;                                           //Разработал
        private string checker;                                             //Проверил
        private string normativeControl;                                    //Нормоконтороль
        private string approver;                                            //Утвердил

        public string Litera { set; get; }                                  //Литера
        public string InvNumbOrigin { set; get; }                           //Инвентарный номер подлинника по ГОСТ 2.501
        public string ReplacedInvNumb { set; get; }                         //Инвентарный номер подлинника, взамен которого выпущен данный подлинник по ГОСТ 2.503
        public string InvNumbDupl { set; get; }                             //Инвентарный номер дубликата по ГОСТ 2.502

        public string ReferenceNumb { set; get; }                           //Обозначение документа, взамен или на основании которого выпущен данный документ
        public string PrimaryApplication { set; get; }                      //Первичное применение

        public string FilePath { set; get; }                                //Путь к файлу текущего элемента
        public string Configuration { set; get; }                           //Конфигурация исходной сборки
        public bool usePdmFlag { set; get; }                                //Флаг, указывающий на принадлежность файла базе PDM

        public DateTime DateDeveloper { set; get; }                 //Дата разработки
        public DateTime DateChecker { set; get; }                   //Дата проверки
        public DateTime DateNormativeControl { set; get; }          //Дата проверки нормоконтролем
        public DateTime DateApprover { set; get; }                  //Дата утверждения

        public string Developer
        {
            set
            {
                if (value != "")
                    DateDeveloper = DateTime.Now;
                developer = value;
            }
            get
            {
                return developer;
            }
        }

        public string Checker
        {
            set
            {
                if (value != "")
                    DateChecker = DateTime.Now;
                checker = value;
            }
            get
            {
                return checker;
            }
        }

        public string NormativeControl
        {
            set
            {
                if (value != "")
                    DateNormativeControl = DateTime.Now;
                normativeControl = value;
            }
            get
            {
                return normativeControl;
            }
        }

        public string Approver
        {
            set
            {
                if (value != "")
                    DateApprover = DateTime.Now;
                approver = value;
            }
            get
            {
                return approver;
            }
        }
    }

}
