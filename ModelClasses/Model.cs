using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VKSMM.ModelClasses
{

    public class Product
    {
        public string IDURL = string.Empty;
        public List<string> URLPhoto = new List<string>();
        public List<string> FilePath = new List<string>();
        public DateTime datePost;
        public int[] prise = new int[5];
        public List<string> sellerText = new List<string>();
        public List<string> sellerTextCleen = new List<string>();
        public List<string> logRegularExpression = new List<string>();
        public string CategoryOfProductName = "ВСЕ";
        public string SubCategoryOfProductName = "ВСЕ";
        public bool HandBlock = false;

        public Guid productGuid = Guid.Empty;
        public int groupIDSendToTelegram = 0;

        public string Materials = "";
        public string Prises = "";
        public string Sizes = "";
    }


    public class CategoryOfProduct
    {
        public CategoryOfProduct()
        {
            SubCategoty.Add(new SubCategoryOfProduct());
        }

        public bool isProvider = true;

        public string Name = "Всякое";
        public List<SubCategoryOfProduct> SubCategoty = new List<SubCategoryOfProduct>();
        public List<Key> Keys = new List<Key>();
    }

    /// <summary>
    /// Класс подкатегории товара
    /// </summary>
    public class SubCategoryOfProduct
    {
        public string Name = "ВСЕ";
        public List<Key> Keys = new List<Key>();
    }

    public class Key
    {
        public bool isProvider = true;
        public bool IsActiv = true;
        public string Value = "";
    }

    public class ColorKeys
    {
        public string Name = "";
        public Color color = Color.White;
        public int Action = -1;
        public Key RegKey = new Key();
    }

    public class ReplaceKeys
    {
        public string Name = "";
        public string NewValue = "";
        public string GroupValue = "ВСЕ";
        public int Action = -1;
        public Key RegKey = new Key();
    }



}
