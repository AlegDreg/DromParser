using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using System;
using System.Data;

namespace DromParser
{
    public class Parser
    {
        IWebDriver web;
        int countItemOnPage = 20;
        int currentPage = 1;
        int countElements;
        
        /// <summary>
        /// Результирующая таблица
        /// </summary>
        public DataTable DataTable { get; private set; }

        /// <summary>
        /// Марка авто
        /// </summary>
        string marka;

        /// <summary>
        /// Модель авто
        /// </summary>
        string models;

        public delegate void Ready();
        /// <summary>
        /// Вызывается после завершения сбора информации
        /// </summary>
        public event Ready ready;

        /// <summary>
        /// Начать сбор
        /// </summary>
        /// <param name="mark">Марка</param>
        /// <param name="model">Модель</param>
        /// <param name="count_elements">Количество собираемых объявлений</param>
        public void Start(string mark, string model, int count_elements)
        {
            marka = mark;
            models = model;

            countElements = count_elements;
            var a = new FirefoxDriver();

            DataTable = new DataTable();
            DataTable.Columns.Add("marka");
            DataTable.Columns.Add("year");
            DataTable.Columns.Add("price");
            DataTable.Columns.Add("probeg");

            web = a;

            OpenNewPage(1);
        }

        void GetAllItemOnPage()
        {
            var a = web.FindElement(By.XPath("/html/body/div[2]/div[4]/div[1]/div[1]/div[4]/div/div[2]"));

            for (int i = 1; i <= countItemOnPage; i++)
            {
                string url = $"/html/body/div[2]/div[4]/div[1]/div[1]/div[4]/div/div[2]/a[{i}]";//элемент

                //var zz = a.FindElement(By.XPath(url));

                string markYear = a.FindElement(By.XPath(url + "/div[2]/div[1]/div/span")).Text;

                string[] asd = markYear.Split(',');

                string mark = asd[0];
                string year = asd[1].Replace(" ", "");


                string zzk = a.FindElement(By.XPath(url + $"/div[2]/div[2]")).Text;//разная инфа включая пробег

                string price = a.FindElement(By.XPath(url + "/div[3]/div[1]/div[1]/span/span")).Text;

                string[] ass = zzk.Split(',');

                int n = -1;

                for (int j = 0; j < ass.Length; j++)
                {
                    string k = ass[j].Replace("тыс", "");
                    if (k != ass[j])
                    {
                        n = j;
                        break;
                    }
                }

                string probeg;

                try
                {
                    string[] azx = ass[n].Split(' ');

                    probeg = azx[1];
                }
                catch
                {
                    continue;
                }

                DataTable.Rows.Add(mark, year, price, probeg);
            }

            if (DataTable.Rows.Count >= countElements)
            {
                ready?.Invoke();

                return;
            }

            OpenNewPage(currentPage + 1);
        }

        void OpenNewPage(int i)
        {
            web.Navigate().GoToUrl($"https://auto.drom.ru/{marka}/{models}/used/page{i}/");
            currentPage = i;

            WebDriverWait wait = new WebDriverWait(web, TimeSpan.FromSeconds(30));
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("/html/body/div[2]/div[4]/div[1]/div[1]/div[4]/div/div[2]/a[1]/div[2]/div[2]")));

            GetAllItemOnPage();
        }
    }
}