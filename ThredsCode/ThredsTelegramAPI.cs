using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Telegram.BotAPI;
using Telegram.BotAPI.AvailableMethods;
using Telegram.BotAPI.AvailableTypes;
using Telegram.BotAPI.GettingUpdates;
using Telegram.BotAPI.UpdatingMessages;
using VKSMM.ModelClasses;
using VKSMM.StuffClasses;

namespace VKSMM.ThredsCode
{
    /// <summary>
    /// Класс с телами основных потоков программы
    /// </summary>
    class ThredsTelegramAPI
    {

        public static void threadBotCode(object mainWindow)
        {
            MainForm mainForm = (MainForm)mainWindow;

            List<Product> productsBuffer = new List<Product>();

            //Name:VKSMMADMINbot
            //API:5313258254:AAE-i9w4hH7RuPVvoTHMqk-dWytW4keeYl8


            var bot = new BotClient("5313258254:AAE-i9w4hH7RuPVvoTHMqk-dWytW4keeYl8");

            //  bot.SetMyCommands(new BotCommand("get", "Собрать посты"));
            //bot.SetMyCommands(new BotCommand("post", "Выложить посты"));
            bot.SetMyCommands(new BotCommand("callback", "new callback"));


            var updates = bot.GetUpdates();
            while (true)
            {

                lock (mainForm.postFromVKs)
                {
                    for (int i = 0; i < mainForm.postFromVKs.Count; i++)
                    {
                        foreach (MainForm.PostOnTheWall post in mainForm.postFromVKs[i].postOnTheWalls)
                        {
                            if (post.sendingOnTelegramm)
                            {
                                break;
                            }
                            else
                            {
                                post.sendingOnTelegramm = true;
                                mainForm.postToTelegramm.Add(post);
                            }
                        }
                    }

                }




                if (updates.Length > 0)
                {
                    foreach (var update in updates)
                    {
                        switch (update.Type)
                        {

                            case UpdateType.Message:
                                var message = update.Message;
                                try
                                {
                                    if (message.Text.Contains("/callback"))
                                    {
                                        if (mainForm.postToTelegramm.Count >= 0)
                                        {

                                            int i = 0;
                                            MainForm.PostOnTheWall postOn;





                                            while (i < 10 && mainForm.postToTelegramm.Count > 0)
                                            {
                                                try
                                                {
                                                    postOn = mainForm.postToTelegramm[0];


                                                    Product P = Stuff.ProcessingProductsForTelegram(mainForm, postOn);
                                                    P.productGuid = Guid.NewGuid();
                                                    productsBuffer.Add(P);


                                                    string resultText = "";
                                                    foreach(string s in P.sellerTextCleen)
                                                    {
                                                        resultText = resultText + s+"\r\n";
                                                    }

                                                    string resoltCat = "Категория: " + P.CategoryOfProductName
                                                        + "\nПодкатегория: " + P.SubCategoryOfProductName;








                                                    List<InputMedia> inputMedias = new List<InputMedia>();
                                                    InputMediaPhoto inputMediaFirst = new InputMediaPhoto();
                                                    inputMediaFirst.Caption = resultText;
                                                    inputMediaFirst.Media = postOn.pictureURL[0];

                                                    inputMedias.Add(inputMediaFirst);

                                                    for (int j = 1; j < postOn.pictureURL.Count; j++)
                                                    {
                                                        inputMediaFirst = new InputMediaPhoto();
                                                        inputMediaFirst.Media = postOn.pictureURL[j];
                                                        inputMedias.Add(inputMediaFirst);
                                                    }


                                                    SendMediaGroupArgs sendMedia = new SendMediaGroupArgs(message.Chat.Id, inputMedias);

                                                    try
                                                    {
                                                        IEnumerable<Message> res = bot.SendMediaGroup(message.Chat.Id, inputMedias, disableNotification: true);

                                                        P.groupIDSendToTelegram = res.ToArray()[0].MessageId;//.MediaGroupId;


                                                        InlineKeyboardButton[] inlineKeyboardButtons = new InlineKeyboardButton[2];
                                                        inlineKeyboardButtons[0] = InlineKeyboardButton.SetCallbackData("Post", "P"+ P.productGuid);
                                                        inlineKeyboardButtons[1] = InlineKeyboardButton.SetCallbackData("Correcting", "C"+ P.productGuid);


                                                        var replyMarkup = new InlineKeyboardMarkup(inlineKeyboardButtons);



                                                        bot.SendMessage(message.Chat.Id, resoltCat, replyMarkup: replyMarkup, disableNotification: true);
                                                    }
                                                    catch
                                                    {

                                                        int k = 0;
                                                        k++;

                                                        //  bot.SendMessage(message.Chat.Id, "ОШИБКА \r\n"+ postOn.text, disableNotification: true);

                                                    }
                                                }
                                                catch
                                                {

                                                }

                                                mainForm.postToTelegramm.RemoveAt(0);

                                                i++;

                                            }



                                        }
                                        else
                                        {
                                            bot.SendMessage(message.Chat.Id, "Новых товаров не обнаружено!", disableNotification: true);
                                        }

                                    }
                                }
                                catch 
                                {
                                }
                                break;


                            case UpdateType.CallbackQuery:
                                try
                                {
                                    var query = update.CallbackQuery;


                                    //bot.AnswerCallbackQuery(query.Id, "RUNING");





                                    bot.EditMessageText(new EditMessageTextArgs
                                    {
                                        ChatId = query.Message.Chat.Id,
                                        MessageId = query.Message.MessageId,
                                        Text = $"OK"
                                        //Text = $"Click!\n\n{query.Data}"
                                    });




                                    Guid g = new Guid(query.Data.Substring(1));


                                    foreach(Product P in productsBuffer)
                                    {
                                        if (P.productGuid == g)
                                        {
                                            int mID = P.groupIDSendToTelegram;
                                            while (mID < query.Message.MessageId)
                                            {
                                                bot.DeleteMessage(query.Message.Chat.Id, mID);
                                                mID++;
                                            }
                                        
                                            break;
                                        }
                                    }


                                }
                                catch
                                {

                                }
                                break;
                        }
                    }
                    updates = updates = bot.GetUpdates(offset: updates.Max(u => u.UpdateId) + 1);
                }
                else
                {
                    updates = bot.GetUpdates();
                }
            }
        }

    }
}
