using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pmi.Service.Interface
{
    interface ICacheService<T>
    {
        bool IsEmpty { get; }
        List<T> GetAll();
        void Add(T item);
        void SetAll(List<T> items);
        void SaveChanges();
        void UploadCache();
    }
}
