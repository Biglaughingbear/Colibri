using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Colibri
{
    class School
    {
        string schoolName;
        Dictionary<int, double> feeDict;
        Dictionary<int, double> amountDict;
        Dictionary<int, double> netDict;
        List<int> orderNums;
        List<Item> itemsList;

        public School(string schoolName)
        {
            this.schoolName = schoolName;
            feeDict = new Dictionary<int, double>();
            amountDict = new Dictionary<int, double>();
            netDict = new Dictionary<int, double>();
            itemsList = new List<Item>();

            orderNums = new List<int>();
        }
        public void AddMoney(int orderNum, double net, double fee, double amount)
        {
            netDict[orderNum]=net;
            feeDict[orderNum]= fee;
            amountDict[orderNum]= amount;
        }
        public void AddItem(string name, int sku, int quant, string size)
        {
            bool original = true;
            foreach (Item item in itemsList)
            {
                if (sku == item.SKU)
                {
                    item.Quant=quant;
                    original = false;
                    break;
                }
            }
            if (original) { 
                itemsList.Add(new Item(name, sku, quant, size));
            }
        }
        public void AddOrderNum(int orderNum)
        {
            orderNums.Add(orderNum);
        }
        public bool HasOrder(int target)
        {
            bool result = false;

            foreach (int num in orderNums)
            {
                if (target==num)
                {
                    result = true;
                }
            }
            return result;
        }
        public double[] GetAllTotals()
        {
            return new double[] { GetTotal(netDict), GetTotal(feeDict), GetTotal(amountDict), feeDict.Keys.Count };
        }
        private double GetTotal(Dictionary<int, double> dict)
        {
            double result = 0;
            foreach (double value in dict.Values)
            {
                result += value;
            }
            return result;
        }
        public Dictionary<int, double[]> GetMoney()
        {
            //feeDict.OrderBy(x=>x.Key);
            Dictionary<int, double[]> results = new Dictionary<int, double[]>();
            foreach (int key in feeDict.Keys)
            {
                results[key] = (new double[] { feeDict[key], netDict[key], amountDict[key] });
            }
            
            return results;
        }
        public List<string[]> GetItems()
        {
            List<string[]> result = new List<string[]>();
            foreach (Item item in itemsList)
            {
                result.Add(new string[] { item.Name, "" + item.Quant, "" + item.SKU });
            }
            return result;
        }
        public List<string[]> GetSortItems()
        {
            List<string[]> result = new List<string[]>();
            foreach (Item item in itemsList.OrderBy(x => x.SKU).ToList())
            {
                result.Add(new string[] { item.Name, "" + item.Quant, "" + item.SKU });
            }
            return result;
        }
    }
    class Item
    {
        string name;
        string size;
        int sku;
        int num;
        int quantity;
        public Item(string name, int sku,int quantity, string size)
        {
            this.name = name;
            this.sku = sku;
            this.quantity = quantity;
            this.num = num;
            this.size = size;
        }

        public string Name
        {
            get
            {
                return name;
            }
        }
        public int SKU
        {
            get
            {
                return sku;
            }
        }
       
        public int Quant
        {
            set
            {
                quantity = quantity + value;
            }
            get
            {
                return quantity;
            }

        }

        public string Size
        {
            get
            {
                return size;
            }
        }
    }
}
