using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace Pmi.Service.Abstraction
{ 
    abstract class CacheService<T>
    {
        protected string filePath;                

        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath">Путь к файлу, в котором будет храниться кэш</param>
        public CacheService(string filePath)
        {
            if (!File.Exists(filePath))
                File.Create(filePath);
            this.filePath = filePath;            
        }           
        /// <summary>
        /// Записывает кэш в файл. Если кэш в файле уже есть, он будет переопределен.
        /// </summary>
        public abstract void Cache(T entity);
        /// <summary>
        /// Подгружает кэш из файла
        /// </summary>
        public abstract T UploadCache();

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
