using Newtonsoft.Json;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace VKSMM.ThredsCode
{
    /// <summary>
    /// Класс с телами потоков работы с сайтом ВКонтакте
    /// </summary>
    class ThredsVKAPI
    {

        //=======================================================================================================
        //Вспомогательные классы для десереализации VKAPI
        //=======================================================================================================

        public class Attachment
        {
            public string type { get; set; }
            public Photo photo { get; set; }
        }

        public class Comments
        {
            public int can_post { get; set; }
            public int count { get; set; }
            public bool groups_can_post { get; set; }
        }

        public class Donut
        {
            public bool is_donut { get; set; }
        }

        public class Item
        {
            public int id { get; set; }
            public int from_id { get; set; }
            public int owner_id { get; set; }
            public int date { get; set; }
            public string post_type { get; set; }
            public string text { get; set; }
            public int is_pinned { get; set; }
            public List<Attachment> attachments { get; set; }
            public PostSource post_source { get; set; }
            public Comments comments { get; set; }
            public Likes likes { get; set; }
            public Reposts reposts { get; set; }
            public Views views { get; set; }
            public Donut donut { get; set; }
            public double short_text_rate { get; set; }
            public int edited { get; set; }
            public string hash { get; set; }
            public int? carousel_offset { get; set; }
        }

        public class Likes
        {
            public int can_like { get; set; }
            public int count { get; set; }
            public int user_likes { get; set; }
            public int can_publish { get; set; }
        }

        public class Photo
        {
            public int album_id { get; set; }
            public int date { get; set; }
            public int id { get; set; }
            public int owner_id { get; set; }
            public string access_key { get; set; }
            public List<Size> sizes { get; set; }
            public string text { get; set; }
            public bool has_tags { get; set; }
            public double? lat { get; set; }
            public double? @long { get; set; }
        }

        public class PostSource
        {
            public string platform { get; set; }
            public string type { get; set; }
        }

        public class Reposts
        {
            public int count { get; set; }
            public int user_reposted { get; set; }
        }

        public class Response
        {
            public int count { get; set; }
            public List<Item> items { get; set; }
        }

        public class Root
        {
            public Response response { get; set; }
        }

        public class Size
        {
            public int height { get; set; }
            public string url { get; set; }
            public string type { get; set; }
            public int width { get; set; }
        }

        public class Views
        {
            public int count { get; set; }
        }
        //=======================================================================================================

        static readonly HttpClient client = new HttpClient();

        /// <summary>
        /// Тело процесса сбора данных с сайта ВКонтакте
        /// </summary>
        public static async void threadVKCollectCode(object mainWindow)
        {

            MainForm mainForm = (MainForm)mainWindow;
            //=======================================================================================================
            Action S0 = () => mainForm.prgBrVKID.Maximum = mainForm.providerVKIDs.Count;
            mainForm.prgBrVKID.Invoke(S0);
            //=======================================================================================================

            

            string wallGetString = "https://api.vk.com/method/wall.get?owner_id=PARAMS&offset=OFFCET&access_token=8320079d8320079d8320079d7e835c90cb883208320079de15d62bd7ffb34813838c72d&v=5.131";
            string line = wallGetString;
            int offcet_count = 1;
            int post_count = 0;
            int postALLCount = 0;
            int iterationToVK = 0;
            int providerNotResponseCount = 0;
            bool eqitPostRead = true;
            DateTime timeLastIteration = DateTime.Now;

            while (true)
            {
                iterationToVK++;
                int i = 0;
                post_count = 0;
                timeLastIteration = DateTime.Now;
                providerNotResponseCount = 0;
                //Для каждого провайдера
                foreach (string idVK in mainForm.providerVKIDs)
                {
                    //=======================================================================================================
                    Action S1 = () => mainForm.prgBrVKID.Value = i;
                    mainForm.prgBrVKID.Invoke(S1);
                    //=======================================================================================================

                    try
                    {

                        offcet_count = 1;
                        eqitPostRead = true;
                        while (eqitPostRead)
                        {

                            line = wallGetString.Replace("PARAMS", idVK);
                            line = line.Replace("OFFCET", offcet_count.ToString());

                            HttpResponseMessage response = await client.GetAsync(line);
                            response.EnsureSuccessStatusCode();
                            string responseBody = await response.Content.ReadAsStringAsync();

                            if (responseBody.Length < 40)
                            {
                                providerNotResponseCount++;
                                break;
                            }

                            //Разбираем JSON полученный от сервера на массив чисел
                            Root numbers = JsonConvert.DeserializeObject<Root>(responseBody);

                            //listBox1.Items.Add(idVK);
                            if (numbers.response != null)
                            {

                                foreach (Item ite in numbers.response.items)
                                {

                                    DateTime date = (new DateTime(1970, 1, 1, 0, 0, 0, 0)).AddSeconds(ite.date);


                                    if (date > DateTime.Now.AddHours(-24) && (ite.attachments != null) && ite.text.Length > 0)
                                    {
                                        MainForm.PostOnTheWall post = new MainForm.PostOnTheWall();
                                        post.datetime = Convert.ToDateTime(date);
                                        post.text = ite.text;
                                        foreach (Attachment kk in ite.attachments)
                                        {
                                            try
                                            {
                                                if (kk.photo != null)
                                                {
                                                    post.pictureURL.Add(kk.photo.sizes.Last().url);
                                                }
                                            }
                                            catch { }
                                        }


                                        lock (mainForm.postFromVKs)
                                        {

                                            bool reg = true;

                                            foreach (MainForm.PostOnTheWall p in mainForm.postFromVKs[i].postOnTheWalls)
                                            {
                                                if (p.text == post.text)
                                                {
                                                    reg = false;
                                                    break;
                                                }
                                            }


                                            if (reg)
                                            {
                                                post_count++;
                                                postALLCount++;
                                                mainForm.postFromVKs[i].postOnTheWalls.Add(post);
                                            }
                                            else
                                            {
                                                eqitPostRead = false;
                                                break;
                                            }
                                        }

                                    }
                                    else
                                    {
                                        eqitPostRead = false;
                                        break;
                                    }
                                }
                            }
                            else
                            {
                                providerNotResponseCount++;
                                break;
                            }

                            offcet_count = offcet_count + 20;

                            Thread.Sleep(400);

                        }
                    }
                    catch { }

                    
                    //=======================================================================================================
                    Action S2 = () => mainForm.lblTelegramIterationInfo.Text = "Процесс опрооса поставщиков: опрошено "+i
                    +" из "+ mainForm.providerVKIDs.Count + "                             Собрано постов: "+ post_count;
                    mainForm.lblTelegramIterationInfo.Invoke(S2);
                    //=======================================================================================================

                    i++;
                }

                //=======================================================================================================
                Action S3 = () => mainForm.txtTelegramResultOneIteration.Text = "\r\nВремя обхода поставщиков:"+(DateTime.Now-timeLastIteration).ToString()
                +"\r\nВсего собрано новых товаров:"+ post_count + "\r\n";
                mainForm.txtTelegramResultOneIteration.Invoke(S3);
                //=======================================================================================================

                //=======================================================================================================
                Action S4 = () => mainForm.txtTelegramResultALLIteration.Text = "\r\nКоличество обходов поставщиков: "+ iterationToVK + "\r\n\r\n"
                + "Всего собрано новых товаров за день:"+ postALLCount + "\r\n\r\n"
                + "Всего доступно поставщиков:"+(i- providerNotResponseCount )+ "\r\n\r\n"
                + "Не доступно поставщиков: "+ providerNotResponseCount + "\r\n\r\n"
                + "Запрошено товаров: 0\r\n\r\n"
                + "Отправлено на пост: 0\r\n\r\n"
                + "Отложено для коррекции: 0\r\n\r\n";
                mainForm.txtTelegramResultALLIteration.Invoke(S4);
                //=======================================================================================================


                Thread.Sleep(1000);

            }
        }

    }
}
