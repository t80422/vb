using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//收集C#成績(五個人)
namespace CsConsole
{
    internal class ArrayForLoop
    {
        public static void Main() 
        {
            //int 為結構類型 沒有參考位址 所以設null會報錯
            //加上"?" 就可以了
            int? value = null;
            //定義陣列 整數類型(一個維度) 沒有考試的給null
            int?[] cs = new int?[5];//若是用 new int[5]則初始值為0
            //指定陣列位址 從0開始
            cs[0] = 100;
            cs[1] = 90;
            cs[3] = 50;
            cs[4] = 30;
            //將每個陣列元素內容(成績)問出來
            for(int i=0;i<cs.Length;i++)
            {
                if (cs[i] == null)
                {
                    Console.WriteLine($"號碼:{i + 1} 缺考");
                }
                else if (cs[i]>=60)
                {
                    Console.WriteLine($"號碼:{i+1} 及格");
                }
                else
                {
                    Console.WriteLine($"號碼:{i + 1} 不及格");
                }
                
            }
        }
    }
}
