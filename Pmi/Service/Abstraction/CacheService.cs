using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace Pmi.Service.Abstraction
{
    /// <summary>
    /// Представляет возможность для работы с кэшем списка однотипных объектов.
    /// </summary>
    /// <typeparam name="T"></typeparam>
    abstract class CacheService<T>
    {
        protected string filePath;
        protected List<T> elements;
        public bool IsEmpty { get { return !elements.Any(); } }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath">Путь к файлу, в котором будет храниться кэш</param>
        public CacheService(string filePath)
        {
            if (!File.Exists(filePath))
            {
                File.WriteAllText(filePath, "[]");
            }
            this.filePath = filePath;
            elements = new List<T>();
        }

        /// <summary>
        /// Возвращает список элементов в кэше
        /// </summary>
        /// <returns></returns>
        public List<T> GetAll()
        {
            return elements;
        }
        /// <summary>
        /// Добавляет элемент в кэш
        /// </summary>
        /// <param name="item"></param>
        public void Add(T item)
        {
            elements.Add(item);
        }
        /// <summary>
        /// Устанавливает список элементов как кэш
        /// </summary>
        /// <param name="items"></param>
        public void SetAll(List<T> items)
        {
            elements = items;
        }
        /// <summary>
        /// Записывает текущий кэш в файл. Если кэш в файле уже есть, он будет переопределен.
        /// </summary>
        public abstract void SaveChanges();
        /// <summary>
        /// Подгружает кэш из файла
        /// </summary>
        public abstract void UploadCache();

        /// <summary>
        /// Вычисляет хэш строки
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        [Obsolete]
        private string GetHash(string value)
        {
            using (var md5 = MD5.Create())
            {
                return Encoding.UTF8.GetString(md5.ComputeHash(Encoding.UTF8.GetBytes(value)));
            }
        }
    }
}
