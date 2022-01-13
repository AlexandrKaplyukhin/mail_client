using OpenPop.Mime;
using OpenPop.Pop3;
using OpenPop.Pop3.Exceptions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Security.AccessControl;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MimeKit;
using MailKit.Security;

namespace Programma
{
    public partial class Form1 : Form
    {
        private static CspParameters cspp = new CspParameters();
        private static RijndaelManaged cryptoRijndael;
        const string keyName = "Key";
        private static RSACryptoServiceProvider rsa;
        private static DSACryptoServiceProvider dsa;
        public static string keyFile = null;
        public static string keyFilee = null;
        public Form1()
        {
            cspp.KeyContainerName = keyName;
            rsa = new RSACryptoServiceProvider(cspp);
            dsa = new DSACryptoServiceProvider();
            rsa.PersistKeyInCsp = true;
            cryptoRijndael = new RijndaelManaged();
            cryptoRijndael.BlockSize = 128;
            cryptoRijndael.KeySize = 128;
            cryptoRijndael.Mode = CipherMode.CBC;
            pop3Client = new Pop3Client();

            imapClient = new ImapClient();
            messages = new Dictionary<int, OpenPop.Mime.Message>(); DirectoryInfo directoryInfo = new DirectoryInfo(Application.StartupPath + "\\Attachments\\");
            DirectoryInfo directorysend = new DirectoryInfo(Application.StartupPath + "\\send\\");
            if (directoryInfo.Exists)
            {
                DirectorySecurity accessControl = directoryInfo.GetAccessControl();
                directoryInfo.Delete(true);
                directoryInfo.Create(accessControl);
            }
            else
                directoryInfo.Create();
            if (directorysend.Exists)
            {
                DirectorySecurity accessControl = directorysend.GetAccessControl();
                directorysend.Delete(true);
                directorysend.Create(accessControl);
            }
            else
                directorysend.Create();
            InitializeComponent();
        }
        //-------------------------Настройки-------------------------
        //private Settings IniFile = new Settings(Application.StartupPath + "\\settings.ini");
        Settings IniFile = new Settings(Application.StartupPath + "\\settings.ini");
        Settings textmessageEmail;
        private Pop3Client pop3Client;
        private ImapClient imapClient;
        private Dictionary<int, OpenPop.Mime.Message> messages;
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            int index = ((TabControl)sender).SelectedIndex;
            switch (index)
            {
                case 0:
                    break;
                case 1:
                    textBox3.Text = IniFile.IniReadValue("send", "address");
                    break;
                case 2:
                    comboBox1.Items.Clear();
                    foreach (string file in Directory.GetFiles(Application.StartupPath, "*.ini"))
                    {
                        FileInfo fi = new FileInfo(file);
                        comboBox1.Items.Add(fi.Name);
                    }
                    comboBox1.SelectedIndex = 0;
                    IniFile = new Settings(Application.StartupPath + "\\" + comboBox1.Text );
                    textBox4.Text = IniFile.IniReadValue("send", "server");
                    textBox5.Text = IniFile.IniReadValue("send", "port");

                    textBox8.Text = IniFile.IniReadValue("recive", "server");
                    textBox9.Text = IniFile.IniReadValue("recive", "port");
                    textBox10.Text = IniFile.IniReadValue("recive", "login");
                    textBox11.Text = IniFile.IniReadValue("recive", "password");
                    checkBox2.Checked = IniFile.IniReadValue("recive", "ssl").Contains("1");
                    checkBox3.Checked = IniFile.IniReadValue("recive", "delmessage").Contains("1");

                    textBox14.Text = IniFile.IniReadValue("recive", "imapserver");
                    textBox15.Text = IniFile.IniReadValue("recive", "imapport");
                    break;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            IniFile.IniWriteValue("send", "login", textBox10.Text);
            IniFile.IniWriteValue("send", "password", textBox11.Text);
            IniFile.IniWriteValue("send", "server", textBox4.Text);
            IniFile.IniWriteValue("send", "port", textBox5.Text);
            IniFile.IniWriteValue("send", "address", textBox10.Text);
            IniFile.IniWriteValue("send", "ssl", checkBox2.Checked ? "1" : "0");

            IniFile.IniWriteValue("recive", "login", textBox10.Text);
            IniFile.IniWriteValue("recive", "password", textBox11.Text);
            IniFile.IniWriteValue("recive", "server", textBox8.Text);
            IniFile.IniWriteValue("recive", "port", textBox9.Text);
            IniFile.IniWriteValue("recive", "ssl", checkBox2.Checked ? "1" : "0");
            IniFile.IniWriteValue("recive", "delmessage", checkBox3.Checked ? "1" : "0");

            IniFile.IniWriteValue("recive", "imapserver", textBox14.Text);
            IniFile.IniWriteValue("recive", "imapport", textBox15.Text);
        }
        //------------------------------входящие сообщения------------------------------
        /// <summary>
        /// Кнопка получения писем
        /// </summary>
        private void button1_Click(object sender, EventArgs e)
        {
            button1.Text = "Обновление...";
            button1.Enabled = false;
            try
            {
                treeView1.Nodes.Clear();
                ReceiveMails();
                SaveMails();
            }
            finally
            {
                button1.Text = "Получить";
                button1.Enabled = true;
            }
        }
        /// <summary>
        /// Метод получения писем с почты
        /// </summary>
        private void ReceiveMails()
        {
            try
            {
                if (pop3Client.Connected)
                    pop3Client.Disconnect();
                int port;
                try
                {
                    port = (int)Convert.ToInt16(IniFile.IniReadValue("recive", "port"));
                }
                catch (Exception)
                {
                    port = 110;
                }
                pop3Client.Connect(IniFile.IniReadValue("recive", "server"), port, IniFile.IniReadValue("recive", "ssl").Contains("1"));
                pop3Client.Authenticate(IniFile.IniReadValue("recive", "login"), IniFile.IniReadValue("recive", "password"));
                int messageCount = pop3Client.GetMessageCount();
                messages.Clear();
                for (int index = 0; index < messageCount; index++)
                //for (int index = messageCount; index >= 1; --index)
                {
                    try
                    {
                        Application.DoEvents();
                        OpenPop.Mime.Message message = pop3Client.GetMessage(index);
                        TreeNode node = new TreeNode(Convert.ToDateTime(message.Headers.Date).ToString("dd-MM-yyyy HH:mm:ss") + " | " + (object)message.Headers.From.MailAddress);
                        node.Tag = (object)index;
                        foreach (OpenPop.Mime.MessagePart allAttachment in message.FindAllAttachments())
                        {
                            if (allAttachment != null)
                            {
                                TreeNode treeNode = node.Nodes.Add(allAttachment.FileName);
                                string str = allAttachment.FileName.Split('.')[1];
                                string publicKey = allAttachment.FileName.Split('.')[0];
                                string folder = "Attachments";
                                if (publicKey == "key")
                                {
                                    folder = "Keys";
                                    FileInfo file = new FileInfo(Application.StartupPath + "\\" + folder + "\\" + allAttachment.FileName);
                                    allAttachment.Save(file);
                                    treeNode.Tag = (object)file;
                                }
                                else
                                {
                                    FileInfo file = new FileInfo(Application.StartupPath + "\\" + folder + "\\" + allAttachment.FileName + "." + DateTime.Now.Ticks.ToString() + "." + str);
                                    allAttachment.Save(file);
                                    treeNode.Tag = (object)file;
                                }
                            }
                        }
                        treeView1.Nodes.Add(node);
                        messages.Add(index, message);
                        if (IniFile.IniReadValue("recive", "delmessage").Contains("1"))
                            pop3Client.DeleteMessage(index);
                    }
                    catch
                    {
                    }
                }
            }
            catch (InvalidLoginException)
            {
                int num = (int)MessageBox.Show((IWin32Window)this, "Сервер не принимает учетные данные пользователя!", "Проверка подлинности сервера POP3");
            }
            catch (PopServerNotFoundException)
            {
                int num = (int)MessageBox.Show((IWin32Window)this, "Сервер не найден", "Получение POP3");
            }
            catch (PopServerLockedException)
            {
                int num = (int)MessageBox.Show((IWin32Window)this, "Доступ к почтовому ящику блокируется", "Учетная запись POP3 заблокирована");
            }
            catch (LoginDelayException)
            {
                int num = (int)MessageBox.Show((IWin32Window)this, "Вход не допускается. Сервер обеспечивает задержку между входами. Вы уже подключились?", "Задержка входа учетной записи POP3");
            }
            catch (Exception ex)
            {
                int num = (int)MessageBox.Show((IWin32Window)this, "Произошла ошибка при получении почты. " + ex.Message, "Получение POP3");
            }
            finally
            {
            }
        }
        /// <summary>
        /// Открытие загруженных атачментов
        /// </summary>
        private void treeView1_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node.Parent == null)
            {
                richTextBox1.Text = Mesbody(e.Node);
                webBrowser1.Visible = false;
            }
            else
            {
                richTextBox1.Text = Mesbody(e.Node.Parent);
                switch ((e.Node.Tag as FileInfo).Extension.ToLowerInvariant())
                {

                    case ".rtf":
                        richTextBox1.Visible = true;
                        richTextBox1.Text = File.ReadAllText((e.Node.Tag as FileInfo).FullName);
                        webBrowser1.Visible = false;
                        break;
                    case ".txt":
                        richTextBox1.Visible = true;
                        richTextBox1.Text = File.ReadAllText((e.Node.Tag as FileInfo).FullName);
                        webBrowser1.Visible = false;
                        break;
                    default:
                        richTextBox1.Visible = false;
                        webBrowser1.Visible = true;
                        webBrowser1.Navigate((e.Node.Tag as FileInfo).FullName);
                        break;
                }
            }
        }
        /// <summary>
        /// Загрузка содержимого сообщений
        /// </summary>
        int i = 0;
        private string Mesbody(TreeNode e)
        {
            string str = "";
            OpenPop.Mime.Message message;
            if (messages.TryGetValue((int)Convert.ToInt16(e.Tag), out message))
            {
                textBox1.Text = message.Headers.From.MailAddress.ToString();
                try
                {
                    label4.Text = message.Headers.To[0].MailAddress.ToString();
                }
                catch { }
                textBox2.Text = message.Headers.Subject;
                OpenPop.Mime.MessagePart plainTextVersion = message.FindFirstPlainTextVersion();
                if (plainTextVersion != null)
                {
                    str = plainTextVersion.GetBodyAsText();
                }
                else
                {
                    List<OpenPop.Mime.MessagePart> allTextVersions = message.FindAllTextVersions();
                    str = allTextVersions.Count < 1 ? "Не удается найти текстовую версию сообщения" : allTextVersions[0].GetBodyAsText();
                }
                textmessageEmail = new Settings(Application.StartupPath + "\\Attachments\\" + "\\" + i+".txt");
                textmessageEmail.IniWriteValue("Письмо", "Кому", label1.Text);
                textmessageEmail.IniWriteValue("Письмо", "От кого", textBox1.Text);
                textmessageEmail.IniWriteValue("Письмо", "Тема", textBox2.Text);
                textmessageEmail.IniWriteValue("Письмо", "Текст сообщения", richTextBox1.Text);
            }
            i++;
            return str;
        }
        //------------------------------отправка сообщений------------------------------
        /// <summary>
        /// Приатачивание файлов
        /// </summary>
        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog
            {
                Multiselect = true,
                Title = "Выберите файлы для отправки",
                InitialDirectory = Application.StartupPath
            };
            dlg.ShowDialog();
            // пользователь вышел из диалога ничего не выбрав
            if (dlg.FileName == String.Empty)
                return;
            foreach (string file in dlg.FileNames)
            {
                File.Copy(file, Application.StartupPath + "\\send\\" + @"\" + Path.GetFileName(file));
            }
        }
        /// <summary>
        /// Удаление приатаченных файлов
        /// </summary>
        private void button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog
            {
                Multiselect = true,
                Title = "Выберите файлы для удаления",
                InitialDirectory = Application.StartupPath + "\\send\\"
            };
            dlg.ShowDialog();
            // пользователь вышел из диалога ничего не выбрав
            if (dlg.FileName == String.Empty)
                return;
            foreach (string file in dlg.FileNames)
            {
                File.Delete(Application.StartupPath + "\\send\\" + @"\" + Path.GetFileName(file));
            }
        }
        /// <summary>
        /// Жирный текст
        /// </summary>
        private void button7_Click(object sender, EventArgs e)
        {
            Clipboard.SetText("<b>" + richTextBox2.SelectedText + "</b>");
            if (richTextBox2.SelectionLength > 0)
            {
                richTextBox2.Paste();
            }
            Clipboard.Clear();
        }
        /// <summary>
        /// Курсив
        /// </summary>
        private void button8_Click(object sender, EventArgs e)
        {
            Clipboard.SetText("<i>" + richTextBox2.SelectedText + "</i>");
            if (richTextBox2.SelectionLength > 0)
            {
                richTextBox2.Paste();
            }
            Clipboard.Clear();
        }
        /// <summary>
        /// Подчёркнутый
        /// </summary>
        private void button9_Click(object sender, EventArgs e)
        {
            Clipboard.SetText("<u>" + richTextBox2.SelectedText + "</u>");
            if (richTextBox2.SelectionLength > 0)
            {
                richTextBox2.Paste();
            }
            Clipboard.Clear();
        }
        /// <summary>
        /// Метод отправки сообщения
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button6_Click(object sender, EventArgs e)
        {
            SendEmail(textBox3.Text, textBox12.Text, textBox13.Text, richTextBox2.Text);
        }
        private void SendEmail(
      string from,
      string to,
      string subject,
      string body)
        {
            MailMessage message = new MailMessage();
            message.From = new MailAddress(from);
            message.To.Add(to);
            message.Subject = subject;
            message.IsBodyHtml = true;
            message.BodyEncoding = Encoding.UTF8;
            message.Body = body;
            Directory.GetFiles("send", "*.*").ToList().ForEach(name => message.Attachments.Add(new Attachment(name, MediaTypeNames.Text.Plain)));
            int port;
            try
            {
                port = (int)Convert.ToInt16(IniFile.IniReadValue("send", "port"));
            }
            catch (Exception)
            {
                port = 25;
            }
            SmtpClient smtpClient = new SmtpClient(IniFile.IniReadValue("send", "server"), port);
            smtpClient.Credentials = (ICredentialsByHost)new NetworkCredential(IniFile.IniReadValue("send", "login"), IniFile.IniReadValue("send", "password"));
            smtpClient.EnableSsl = IniFile.IniReadValue("send", "ssl").Contains("1");
            button6.Text = "Ожидание...";
            button6.Enabled = false;
            try
            {
                smtpClient.Send(message);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                button6.Text = "Отправить";
                button6.Enabled = true;
            }
        }
        //------------------------------шифрование сообщений------------------------------
        /// <summary>
        /// Кнопка для расшифровки сообщения
        /// </summary>
        private void DecryptFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog()
            {
                InitialDirectory = Application.StartupPath + "\\Attachments\\",
                RestoreDirectory = false
            };

            openFile.Title = "Выберите файл для расшифрования.";
            if (openFile.ShowDialog() == DialogResult.OK)
            {
                string fileName = openFile.FileName;
                if (fileName != null)
                {
                    SaveFileDialog saveFile = new SaveFileDialog()
                    {
                        InitialDirectory = Application.StartupPath + "\\send\\",
                        RestoreDirectory = false
                    };

                    saveFile.Filter = "Шифр (*.txt)|*.txt|All files (*.*)|*.*";

                    saveFile.Title = "Укажите имя для расшифрованного файла.";
                    if (saveFile.ShowDialog() == DialogResult.OK)
                    {
                        string encryptFileName = saveFile.FileName/* + ".txt"*/;
                        if (encryptFileName != null)
                        {
                            Decrypt(fileName, encryptFileName);
                        }
                    }
                }
            }
        }
        /// <summary>
        /// Метод для расшифровки сообщения
        /// </summary>
        private static void Decrypt(string fileName, string encryptFileName)
        {
            FileStream inFs = new FileStream(fileName, FileMode.Open);

            byte[] LenK = new byte[4];
            inFs.Read(LenK, 0, 4);
            byte[] LenIV = new byte[4];
            inFs.Read(LenIV, 0, 4);

            int lenK = BitConverter.ToInt32(LenK, 0);
            int lenIV = BitConverter.ToInt32(LenIV, 0);

            int startC = lenK + lenIV + 8;
            int lenC = (int)inFs.Length - startC;

            byte[] KeyEncrypted = new byte[lenK];
            byte[] IV = new byte[lenIV];

            inFs.Seek(8, SeekOrigin.Begin);
            inFs.Read(KeyEncrypted, 0, lenK);
            inFs.Seek(8 + lenK, SeekOrigin.Begin);
            inFs.Read(IV, 0, lenIV);

            byte[] KeyDecrypted = rsa.Decrypt(KeyEncrypted, false);

            ICryptoTransform transform = cryptoRijndael.CreateDecryptor(KeyDecrypted, IV);

            FileStream outFs = new FileStream(encryptFileName, FileMode.Create);

            int count = 0;
            int offset = 0;

            int blockSizeBytes = cryptoRijndael.BlockSize / 8;
            byte[] data = new byte[blockSizeBytes];


            inFs.Seek(startC, SeekOrigin.Begin);
            CryptoStream outStreamDecrypted = new CryptoStream(outFs, transform, CryptoStreamMode.Write);
            do
            {
                count = inFs.Read(data, 0, blockSizeBytes);
                offset += count;
                outStreamDecrypted.Write(data, 0, count);
            }
            while (count > 0);

            outStreamDecrypted.FlushFinalBlock();
            outStreamDecrypted.Close();
            outFs.Close();
            inFs.Close();

        }
        /// <summary>
        /// Экспорт ключа для текста
        /// </summary>
        private void ExportPublicKey_Click(object sender, EventArgs e)
        {
            try
            {
                keyFile = IniFile.IniReadValue("send", "address");
                StreamWriter sw = new StreamWriter("send\\key." + keyFile + ".key");
                sw.Write(rsa.ToXmlString(false));
                sw.Close();
                label21.Text = "Публичный ключ экспортирован в папку";
            }
            catch (Exception)
            {
                MessageBox.Show("Что-то пошло не так. Попробуйте ещё раз.");
            }
        }
        /// <summary>
        /// Импорт ключа для текста
        /// </summary>
        private void ImportPublicKey_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog()
            {
                InitialDirectory = Application.StartupPath + "\\Keys\\",
                RestoreDirectory = false
            };
            openFile.Title = "Введите файл для импорта открытого ключа";
            if (openFile.ShowDialog() == DialogResult.OK)
            {
                keyFile = openFile.FileName;
            }
            label22.Text = "Публичный ключ импортирован";
        }
        /// <summary>
        /// Кнопка шифрования файла
        /// </summary>
        private void EncryptFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog()
            {
                InitialDirectory = Application.StartupPath,
                RestoreDirectory = false
            };

            openFile.Title = "Выберите файл для шифрования.";
            if (openFile.ShowDialog() == DialogResult.OK)
            {
                string fileName = openFile.FileName;
                if (fileName != null)
                {
                    SaveFileDialog saveFile = new SaveFileDialog()
                    {
                        InitialDirectory = Application.StartupPath + "\\send\\",
                        RestoreDirectory = false
                    };

                    saveFile.Filter = "Шифр (*.txt)|*.txt|All files (*.*)|*.*";

                    saveFile.Title = "Укажите имя для зашифрованного файла.";
                    if (saveFile.ShowDialog() == DialogResult.OK)
                    {
                        string encryptFileName = saveFile.FileName/* + ".txt"*/;
                        if (encryptFileName != null)
                        {
                            Encrypt(fileName, encryptFileName, keyFile);
                        }
                    }
                }
            }
        }
        /// <summary>
        /// Метод для шифрования файла
        /// </summary>
        private static void Encrypt(string fileName, string encryptFileName, string keyFile)
        {
            if (keyFile != null)
            {
                StreamReader sr = new StreamReader(keyFile);
                rsa.FromXmlString(sr.ReadToEnd());
                sr.Close();
            }
            ICryptoTransform transform = cryptoRijndael.CreateEncryptor();
            byte[] keyEncrypted = rsa.Encrypt(cryptoRijndael.Key, false);

            int lKey = keyEncrypted.Length;
            byte[] LenK = BitConverter.GetBytes(lKey);
            int lIV = cryptoRijndael.IV.Length;
            byte[] LenIV = BitConverter.GetBytes(lIV);

            FileStream outFs = new FileStream(encryptFileName, FileMode.Create);
            outFs.Write(LenK, 0, 4);
            outFs.Write(LenIV, 0, 4);
            outFs.Write(keyEncrypted, 0, lKey);
            outFs.Write(cryptoRijndael.IV, 0, lIV);

            CryptoStream outStreamEncrypted = new CryptoStream(outFs, transform, CryptoStreamMode.Write);
            int count = 0;
            int offset = 0;

            // blockSizeBytes can be any arbitrary size.
            int blockSize = cryptoRijndael.BlockSize / 8;
            byte[] data = new byte[blockSize];
            int bytesRead = 0;

            FileStream inFs = new FileStream(fileName, FileMode.Open);
            do
            {
                count = inFs.Read(data, 0, blockSize);
                offset += count;
                outStreamEncrypted.Write(data, 0, count);
                bytesRead += blockSize;
            }
            while (count > 0);
            inFs.Close();
            outStreamEncrypted.FlushFinalBlock();
            outStreamEncrypted.Close();
        }
        /// <summary>
        /// Кнопка для выполнения цифровой подписи
        /// </summary>
        private void button11_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Title = "Выберите необходимый файл.";
            if (openFile.ShowDialog() == DialogResult.OK)
            {
                string fileName = openFile.FileName;
                if (fileName != null)
                {
                    SaveFileDialog saveDialog = new SaveFileDialog()
                    {
                        InitialDirectory = Application.StartupPath + "\\send\\",
                        RestoreDirectory = false
                    };
                    saveDialog.Title = "Укажите имя для подписанного файла.";
                    if (saveDialog.ShowDialog() == DialogResult.OK)
                    {
                        string newFile = saveDialog.FileName + ".txt";

                        if (newFile != null)
                        {
                            Podpis(fileName, newFile);
                        }
                    }
                }
            }
        }
        /// <summary>
        /// Получение хеша текста
        /// </summary>
        public static byte[] GetHash(byte[] bytes)
        {
            using (var sha = SHA1.Create())
            {
                var hash = sha.ComputeHash(sha.ComputeHash(bytes));
                return hash;
            }
        }

        /// <summary>
        /// Метод для выполенения ЭЦП
        /// </summary>
        private static void Podpis(string fileName, string newFile)
        {
            byte[] data = File.ReadAllBytes(fileName);

            DSASignatureFormatter dsaFormatter = new DSASignatureFormatter(dsa);

            dsaFormatter.SetHashAlgorithm("SHA1");

            byte[] signature = dsaFormatter.CreateSignature(GetHash(data));

            int lSign = signature.Length;

            byte[] LenSign = BitConverter.GetBytes(lSign);
            using (FileStream outFs = new FileStream(newFile, FileMode.Create))
            {
                outFs.Write(LenSign, 0, 4);
                outFs.Write(signature, 0, lSign);
                outFs.Write(data, 0, data.Length);
            }
        }
        /// <summary>
        /// Кнопка для проверки цифровой подписи
        /// </summary>
        private void button10_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog()
            {
                InitialDirectory = Application.StartupPath + "\\Attachments\\",
                RestoreDirectory = false
            };
            openFile.Title = "Выберите файл для проверки цифровой подписи.";
            if (openFile.ShowDialog() == DialogResult.OK)
            {
                string fileName = openFile.FileName;
                if (fileName != null)
                {
                    SaveFileDialog saveFile = new SaveFileDialog();
                    saveFile.Title = "Введите путь к файлу для сохранения данных";
                    if (saveFile.ShowDialog() == DialogResult.OK)
                    {
                        string newFile = saveFile.FileName + ".txt";

                        if (newFile != null)
                        {
                            bool success = ProverkaPodpisi(fileName, newFile, keyFilee);
                            if (success)
                            {
                                MessageBox.Show("Проверка подписи пройдена!");
                            }
                            else
                            {
                                MessageBox.Show("Проверка подписи не пройдена!");
                            }
                        }
                    }
                }
            }
        }
        /// <summary>
        /// Метод для проверки ЭЦП
        /// </summary>
        private static bool ProverkaPodpisi(string fileName, string newFile, string keyFilee)
        {
            if (keyFilee != null)
            {
                StreamReader sr = new StreamReader("send\\key." + keyFilee + ".keyDSA");
                dsa.FromXmlString(sr.ReadToEnd());
                sr.Close();
            }

            using (FileStream inFs = new FileStream(fileName, FileMode.Open))
            {
                try
                {
                    byte[] LenSign = new byte[4];
                    inFs.Read(LenSign, 0, 4);

                    int lenSign = BitConverter.ToInt32(LenSign, 0);

                    int startData = lenSign + 4;
                    int lenData = (int)inFs.Length - startData;

                    byte[] sign = new byte[lenSign];
                    byte[] data = new byte[lenData];

                    inFs.Read(sign, 0, lenSign);
                    inFs.Read(data, 0, lenData);

                    File.WriteAllBytes(newFile, data);
                    return true;
                }
                catch (Exception e) { return false; }

            }
        }
        /// <summary>
        /// Экспорт ключа для ЭЦП
        /// </summary>
        private void button12_Click(object sender, EventArgs e)
        {
            try
            {
                keyFilee = IniFile.IniReadValue("send", "address");
                StreamWriter sw = new StreamWriter("send\\key." + keyFilee + ".keyDSA");
                sw.Write(dsa.ToXmlString(false));
                sw.Close();
            }
            catch (Exception)
            {
                MessageBox.Show("Что-то пошло не так. Попробуйте ещё раз.");
            }
            label23.Text = "Публичный ключ экспортирован";
        }

        /// <summary>
        /// Импорт ключа для ЭЦП
        /// </summary>
        private void button13_Click(object sender, EventArgs e)
        {

            OpenFileDialog openFile = new OpenFileDialog()
            {
                InitialDirectory = Application.StartupPath + "\\Keys\\",
                RestoreDirectory = false
            };
            openFile.Title = "Введите файл для импорта открытого ключа";
            if (openFile.ShowDialog() == DialogResult.OK)
            {
                keyFilee = openFile.FileName;
            }
            label24.Text = "Публичный ключ импортирован";
        }


        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            IniFile = new Settings(Application.StartupPath + "\\" + comboBox1.Text);
            textBox4.Text = IniFile.IniReadValue("send", "server");
            textBox5.Text = IniFile.IniReadValue("send", "port");

            textBox8.Text = IniFile.IniReadValue("recive", "server");
            textBox9.Text = IniFile.IniReadValue("recive", "port");
            textBox10.Text = IniFile.IniReadValue("recive", "login");
            textBox11.Text = IniFile.IniReadValue("recive", "password");
            checkBox2.Checked = IniFile.IniReadValue("recive", "ssl").Contains("1");
            checkBox3.Checked = IniFile.IniReadValue("recive", "delmessage").Contains("1");
            i = 0;
        }
        bool proverka = false;
        private void button14_Click(object sender, EventArgs e)
        {
            if (proverka == false)
            {
                textBox4.Text = ""; textBox5.Text = ""; textBox14.Text = ""; textBox15.Text = "";
                textBox8.Text = ""; textBox9.Text = ""; textBox10.Text = ""; textBox11.Text = "";
                proverka = true;
            }
            else
            {
                IniFile = new Settings(Application.StartupPath + "\\" + textBox10.Text + ".ini");

                IniFile.IniWriteValue("send", "login", textBox10.Text);
                IniFile.IniWriteValue("send", "password", textBox11.Text);
                IniFile.IniWriteValue("send", "server", textBox4.Text);
                IniFile.IniWriteValue("send", "port", textBox5.Text);
                IniFile.IniWriteValue("send", "address", textBox10.Text);

                IniFile.IniWriteValue("recive", "login", textBox10.Text);
                IniFile.IniWriteValue("recive", "password", textBox11.Text);
                IniFile.IniWriteValue("recive", "server", textBox8.Text);
                IniFile.IniWriteValue("recive", "port", textBox9.Text);
                IniFile.IniWriteValue("recive", "ssl", checkBox2.Checked ? "1" : "0");
                IniFile.IniWriteValue("recive", "delmessage", checkBox3.Checked ? "1" : "0");

                IniFile.IniWriteValue("recive", "imapserver", textBox14.Text);
                IniFile.IniWriteValue("recive", "imapport", textBox15.Text);

                proverka = false;
                MessageBox.Show("Добавлена учётная запись");

                comboBox1.Items.Clear();
                foreach (string file in Directory.GetFiles(Application.StartupPath, "*.ini"))
                {
                    FileInfo fi = new FileInfo(file);
                    comboBox1.Items.Add(fi.Name);
                }
                comboBox1.SelectedIndex = 0;
            }
        }
        /// <summary>
        /// Сохранение всех сообщений
        /// </summary>
        string Subject = "";
        string From = "";
        string To = "";
        string Cc = "";
        string Sender = "";
        string TextBody = "";
        int j = 0;
        string folder = "";
        string mesage = "";
        int count = 0;
        private void SaveMails()
        {
            for (int papka = 0; papka < 5; papka++)
            {
                try
                {
                    j = 0;
                    string imap = "imap.yandex.ru"; // Получаем imap сервер из конфига по полю Login
                    using (var client = new ImapClient())
                    {
                        client.Connect(imap, 993, true); // Конект к серверу IMap первое значение это название сервера,второе это порт сервера и третье используется ли SSL подключение
                        client.AuthenticationMechanisms.Remove("XOAUTH");
                        string login = IniFile.IniReadValue("send", "login");
                        string password = IniFile.IniReadValue("send", "password");
                        client.Authenticate(login, password); // Авторизируемся данными из полей
                        if (papka == 0) 
                        { 
                            client.Inbox.Open(FolderAccess.ReadWrite);
                            count = client.Inbox.Count;
                        }
                        if (papka == 1) { 
                            client.GetFolder(SpecialFolder.Trash).Open(FolderAccess.ReadWrite);
                            count = client.GetFolder(SpecialFolder.Trash).Count;
                        }
                        if (papka == 2) { 
                            client.GetFolder(SpecialFolder.Sent).Open(FolderAccess.ReadWrite);
                            count = client.GetFolder(SpecialFolder.Sent).Count;
                        }
                        if (papka == 3) {
                            client.GetFolder(SpecialFolder.Drafts).Open(FolderAccess.ReadWrite);
                            count = client.GetFolder(SpecialFolder.Drafts).Count;
                        }
                        if (papka == 4) { 
                            client.GetFolder(SpecialFolder.Junk).Open(FolderAccess.ReadWrite);
                            count = client.GetFolder(SpecialFolder.Junk).Count;
                        }
                        if (j < count)
                        {
                            if (papka == 0)
                            {
                                var items = client.Inbox.Fetch(0, -1, MessageSummaryItems.UniqueId | MessageSummaryItems.BodyStructure);
                                foreach (var item in items)
                                {
                                    var message = client.Inbox.GetMessage(j);
                                    ////////////////////////////
                                    if (message.Subject != null)
                                    {
                                        Subject = "Subject: " + message.Subject.ToString() + "\n";
                                    }
                                    else
                                    {
                                        Subject = "Subject: " + "Пусто" + "\n";
                                    }
                                    ////////////////////////////
                                    if (message.From != null)
                                    {
                                        From = "From: " + message.From.ToString() + "\n";
                                    }
                                    else
                                    {
                                        From = "From: " + "Пусто" + "\n";
                                    }
                                    ////////////////////////////
                                    if (message.To != null)
                                    {
                                        To = "To: " + message.To.ToString() + "\n";
                                    }
                                    else
                                    {
                                        To = "To: " + "Пусто" + "\n";
                                    }
                                    ////////////////////////////
                                    if (message.Cc != null)
                                    {
                                        Cc = "Cc: " + message.Cc.ToString() + "\n";
                                    }
                                    else
                                    {
                                        Cc = "Cc: " + "Пусто" + "\n";
                                    }
                                    ////////////////////////////
                                    if (message.Sender != null)
                                    {
                                        Sender = "Sender: " + message.Sender.ToString() + "\n";
                                    }
                                    else
                                    {
                                        Sender = "Sender: " + "Пусто" + "\n";
                                    }
                                    ////////////////////////////
                                    if (message.TextBody != null)
                                    {
                                        TextBody = "TextBody: " + message.TextBody.ToString() + "\n";
                                    }
                                    if (message.HtmlBody != null)
                                    {
                                        TextBody = "TextBody: " + message.HtmlBody.ToString() + "\n";
                                    }
                                    else
                                    {
                                        TextBody = "TextBody: " + "Не удается найти текстовую версию сообщения" + "\n";
                                    }
                                    var directory = Path.Combine(Application.StartupPath + "\\inbox\\", item.UniqueId.ToString());
                                    Directory.CreateDirectory(directory);
                                    using (FileStream fstream = new FileStream(directory + "\\Message.txt", FileMode.OpenOrCreate))
                                    {
                                        byte[] input = Encoding.Default.GetBytes(Sender + From + To + Cc + Subject + TextBody);
                                        fstream.Write(input, 0, input.Length);

                                    }
                                    foreach (var attachment in item.Attachments)
                                    {
                                        // download the attachment just like we did with the body
                                        var entity = client.Inbox.GetBodyPart(item.UniqueId, attachment);

                                        // attachments can be either message/rfc822 parts or regular MIME parts
                                        if (entity is MimeKit.MessagePart)
                                        {
                                            var rfc822 = (MimeKit.MessagePart)entity;

                                            var path = Path.Combine(directory, attachment.PartSpecifier + ".eml");

                                            rfc822.Message.WriteTo(path);
                                        }
                                        else
                                        {
                                            var part = (MimePart)entity;

                                            // note: it's possible for this to be null, but most will specify a filename
                                            var fileName = part.FileName;

                                            var path = Path.Combine(directory, fileName);

                                            // decode and save the content to a file
                                            using (var stream = File.Create(path))
                                                part.Content.DecodeTo(stream);
                                        }
                                    }
                                    j++;
                                }
                            }
                            if (papka == 1)
                            {
                                var items = client.GetFolder(SpecialFolder.Trash).Fetch(0, -1, MessageSummaryItems.UniqueId | MessageSummaryItems.BodyStructure);
                                foreach (var item in items)
                                {
                                    var message = client.GetFolder(SpecialFolder.Trash).GetMessage(j);
                                    ////////////////////////////
                                    if (message.Subject != null)
                                    {
                                        Subject = "Subject: " + message.Subject.ToString() + "\n";
                                    }
                                    else
                                    {
                                        Subject = "Subject: " + "Пусто" + "\n";
                                    }
                                    ////////////////////////////
                                    if (message.From != null)
                                    {
                                        From = "From: " + message.From.ToString() + "\n";
                                    }
                                    else
                                    {
                                        From = "From: " + "Пусто" + "\n";
                                    }
                                    ////////////////////////////
                                    if (message.To != null)
                                    {
                                        To = "To: " + message.To.ToString() + "\n";
                                    }
                                    else
                                    {
                                        To = "To: " + "Пусто" + "\n";
                                    }
                                    ////////////////////////////
                                    if (message.Cc != null)
                                    {
                                        Cc = "Cc: " + message.Cc.ToString() + "\n";
                                    }
                                    else
                                    {
                                        Cc = "Cc: " + "Пусто" + "\n";
                                    }
                                    ////////////////////////////
                                    if (message.Sender != null)
                                    {
                                        Sender = "Sender: " + message.Sender.ToString() + "\n";
                                    }
                                    else
                                    {
                                        Sender = "Sender: " + "Пусто" + "\n";
                                    }
                                    ////////////////////////////
                                    if (message.TextBody != null)
                                    {
                                        TextBody = "TextBody: " + message.TextBody.ToString() + "\n";
                                    }
                                    if (message.HtmlBody != null)
                                    {
                                        TextBody = "TextBody: " + message.HtmlBody.ToString() + "\n";
                                    }
                                    else
                                    {
                                        TextBody = "TextBody: " + "Не удается найти текстовую версию сообщения" + "\n";
                                    }
                                    var directory = Path.Combine(Application.StartupPath + "\\trash\\", item.UniqueId.ToString());
                                    Directory.CreateDirectory(directory);
                                    using (FileStream fstream = new FileStream(directory + "\\Message.txt", FileMode.OpenOrCreate))
                                    {
                                        byte[] input = Encoding.Default.GetBytes(Sender + From + To + Cc + Subject + TextBody);
                                        fstream.Write(input, 0, input.Length);

                                    }
                                    foreach (var attachment in item.Attachments)
                                    {
                                        // download the attachment just like we did with the body
                                        var entity = client.GetFolder(SpecialFolder.Trash).GetBodyPart(item.UniqueId, attachment);

                                        // attachments can be either message/rfc822 parts or regular MIME parts
                                        if (entity is MimeKit.MessagePart)
                                        {
                                            var rfc822 = (MimeKit.MessagePart)entity;

                                            var path = Path.Combine(directory, attachment.PartSpecifier + ".eml");

                                            rfc822.Message.WriteTo(path);
                                        }
                                        else
                                        {
                                            var part = (MimePart)entity;

                                            // note: it's possible for this to be null, but most will specify a filename
                                            var fileName = part.FileName;

                                            var path = Path.Combine(directory, fileName);

                                            // decode and save the content to a file
                                            using (var stream = File.Create(path))
                                                part.Content.DecodeTo(stream);
                                        }
                                    }
                                    j++;
                                }
                            }
                            if (papka == 2)
                            {
                                var items = client.GetFolder(SpecialFolder.Sent).Fetch(0, -1, MessageSummaryItems.UniqueId | MessageSummaryItems.BodyStructure);
                                foreach (var item in items)
                                {
                                    var message = client.GetFolder(SpecialFolder.Sent).GetMessage(j);
                                    ////////////////////////////
                                    if (message.Subject != null)
                                    {
                                        Subject = "Subject: " + message.Subject.ToString() + "\n";
                                    }
                                    else
                                    {
                                        Subject = "Subject: " + "Пусто" + "\n";
                                    }
                                    ////////////////////////////
                                    if (message.From != null)
                                    {
                                        From = "From: " + message.From.ToString() + "\n";
                                    }
                                    else
                                    {
                                        From = "From: " + "Пусто" + "\n";
                                    }
                                    ////////////////////////////
                                    if (message.To != null)
                                    {
                                        To = "To: " + message.To.ToString() + "\n";
                                    }
                                    else
                                    {
                                        To = "To: " + "Пусто" + "\n";
                                    }
                                    ////////////////////////////
                                    if (message.Cc != null)
                                    {
                                        Cc = "Cc: " + message.Cc.ToString() + "\n";
                                    }
                                    else
                                    {
                                        Cc = "Cc: " + "Пусто" + "\n";
                                    }
                                    ////////////////////////////
                                    if (message.Sender != null)
                                    {
                                        Sender = "Sender: " + message.Sender.ToString() + "\n";
                                    }
                                    else
                                    {
                                        Sender = "Sender: " + "Пусто" + "\n";
                                    }
                                    ////////////////////////////
                                    if (message.TextBody != null)
                                    {
                                        TextBody = "TextBody: " + message.TextBody.ToString() + "\n";
                                    }
                                    if (message.HtmlBody != null)
                                    {
                                        TextBody = "TextBody: " + message.HtmlBody.ToString() + "\n";
                                    }
                                    else
                                    {
                                        TextBody = "TextBody: " + "Не удается найти текстовую версию сообщения" + "\n";
                                    }
                                    var directory = Path.Combine(Application.StartupPath + "\\sent\\", item.UniqueId.ToString());
                                    Directory.CreateDirectory(directory);
                                    using (FileStream fstream = new FileStream(directory + "\\Message.txt", FileMode.OpenOrCreate))
                                    {
                                        byte[] input = Encoding.Default.GetBytes(Sender + From + To + Cc + Subject + TextBody);
                                        fstream.Write(input, 0, input.Length);

                                    }
                                    foreach (var attachment in item.Attachments)
                                    {
                                        // download the attachment just like we did with the body
                                        var entity = client.GetFolder(SpecialFolder.Sent).GetBodyPart(item.UniqueId, attachment);

                                        // attachments can be either message/rfc822 parts or regular MIME parts
                                        if (entity is MimeKit.MessagePart)
                                        {
                                            var rfc822 = (MimeKit.MessagePart)entity;

                                            var path = Path.Combine(directory, attachment.PartSpecifier + ".eml");

                                            rfc822.Message.WriteTo(path);
                                        }
                                        else
                                        {
                                            var part = (MimePart)entity;

                                            // note: it's possible for this to be null, but most will specify a filename
                                            var fileName = part.FileName;

                                            var path = Path.Combine(directory, fileName);

                                            // decode and save the content to a file
                                            using (var stream = File.Create(path))
                                                part.Content.DecodeTo(stream);
                                        }
                                    }
                                    j++;
                                }
                            }
                            if (papka == 3)
                            {
                                var items = client.GetFolder(SpecialFolder.Drafts).Fetch(0, -1, MessageSummaryItems.UniqueId | MessageSummaryItems.BodyStructure);
                                foreach (var item in items)
                                {
                                    var message = client.GetFolder(SpecialFolder.Drafts).GetMessage(j);
                                    ////////////////////////////
                                    if (message.Subject != null)
                                    {
                                        Subject = "Subject: " + message.Subject.ToString() + "\n";
                                    }
                                    else
                                    {
                                        Subject = "Subject: " + "Пусто" + "\n";
                                    }
                                    ////////////////////////////
                                    if (message.From != null)
                                    {
                                        From = "From: " + message.From.ToString() + "\n";
                                    }
                                    else
                                    {
                                        From = "From: " + "Пусто" + "\n";
                                    }
                                    ////////////////////////////
                                    if (message.To != null)
                                    {
                                        To = "To: " + message.To.ToString() + "\n";
                                    }
                                    else
                                    {
                                        To = "To: " + "Пусто" + "\n";
                                    }
                                    ////////////////////////////
                                    if (message.Cc != null)
                                    {
                                        Cc = "Cc: " + message.Cc.ToString() + "\n";
                                    }
                                    else
                                    {
                                        Cc = "Cc: " + "Пусто" + "\n";
                                    }
                                    ////////////////////////////
                                    if (message.Sender != null)
                                    {
                                        Sender = "Sender: " + message.Sender.ToString() + "\n";
                                    }
                                    else
                                    {
                                        Sender = "Sender: " + "Пусто" + "\n";
                                    }
                                    ////////////////////////////
                                    if (message.TextBody != null)
                                    {
                                        TextBody = "TextBody: " + message.TextBody.ToString() + "\n";
                                    }
                                    if (message.HtmlBody != null)
                                    {
                                        TextBody = "TextBody: " + message.HtmlBody.ToString() + "\n";
                                    }
                                    else
                                    {
                                        TextBody = "TextBody: " + "Не удается найти текстовую версию сообщения" + "\n";
                                    }
                                    var directory = Path.Combine(Application.StartupPath + "\\drafts\\", item.UniqueId.ToString());
                                    Directory.CreateDirectory(directory);
                                    using (FileStream fstream = new FileStream(directory + "\\Message.txt", FileMode.OpenOrCreate))
                                    {
                                        byte[] input = Encoding.Default.GetBytes(Sender + From + To + Cc + Subject + TextBody);
                                        fstream.Write(input, 0, input.Length);

                                    }
                                    foreach (var attachment in item.Attachments)
                                    {
                                        // download the attachment just like we did with the body
                                        var entity = client.GetFolder(SpecialFolder.Drafts).GetBodyPart(item.UniqueId, attachment);

                                        // attachments can be either message/rfc822 parts or regular MIME parts
                                        if (entity is MimeKit.MessagePart)
                                        {
                                            var rfc822 = (MimeKit.MessagePart)entity;

                                            var path = Path.Combine(directory, attachment.PartSpecifier + ".eml");

                                            rfc822.Message.WriteTo(path);
                                        }
                                        else
                                        {
                                            var part = (MimePart)entity;

                                            // note: it's possible for this to be null, but most will specify a filename
                                            var fileName = part.FileName;

                                            var path = Path.Combine(directory, fileName);

                                            // decode and save the content to a file
                                            using (var stream = File.Create(path))
                                                part.Content.DecodeTo(stream);
                                        }
                                    }
                                    j++;
                                }
                            }
                            if (papka == 4)
                            {
                                var items = client.GetFolder(SpecialFolder.Junk).Fetch(0, -1, MessageSummaryItems.UniqueId | MessageSummaryItems.BodyStructure);
                                foreach (var item in items)
                                {
                                    var message = client.GetFolder(SpecialFolder.Junk).GetMessage(j);
                                    ////////////////////////////
                                    if (message.Subject != null)
                                    {
                                        Subject = "Subject: " + message.Subject.ToString() + "\n";
                                    }
                                    else
                                    {
                                        Subject = "Subject: " + "Пусто" + "\n";
                                    }
                                    ////////////////////////////
                                    if (message.From != null)
                                    {
                                        From = "From: " + message.From.ToString() + "\n";
                                    }
                                    else
                                    {
                                        From = "From: " + "Пусто" + "\n";
                                    }
                                    ////////////////////////////
                                    if (message.To != null)
                                    {
                                        To = "To: " + message.To.ToString() + "\n";
                                    }
                                    else
                                    {
                                        To = "To: " + "Пусто" + "\n";
                                    }
                                    ////////////////////////////
                                    if (message.Cc != null)
                                    {
                                        Cc = "Cc: " + message.Cc.ToString() + "\n";
                                    }
                                    else
                                    {
                                        Cc = "Cc: " + "Пусто" + "\n";
                                    }
                                    ////////////////////////////
                                    if (message.Sender != null)
                                    {
                                        Sender = "Sender: " + message.Sender.ToString() + "\n";
                                    }
                                    else
                                    {
                                        Sender = "Sender: " + "Пусто" + "\n";
                                    }
                                    ////////////////////////////
                                    if (message.TextBody != null)
                                    {
                                        TextBody = "TextBody: " + message.TextBody.ToString() + "\n";
                                    }
                                    if (message.HtmlBody != null)
                                    {
                                        TextBody = "TextBody: " + message.HtmlBody.ToString() + "\n";
                                    }
                                    else
                                    {
                                        TextBody = "TextBody: " + "Не удается найти текстовую версию сообщения" + "\n";
                                    }
                                    var directory = Path.Combine(Application.StartupPath + "\\junk\\", item.UniqueId.ToString());
                                    Directory.CreateDirectory(directory);
                                    using (FileStream fstream = new FileStream(directory + "\\Message.txt", FileMode.OpenOrCreate))
                                    {
                                        byte[] input = Encoding.Default.GetBytes(Sender + From + To + Cc + Subject + TextBody);
                                        fstream.Write(input, 0, input.Length);

                                    }
                                    foreach (var attachment in item.Attachments)
                                    {
                                        // download the attachment just like we did with the body
                                        var entity = client.GetFolder(SpecialFolder.Junk).GetBodyPart(item.UniqueId, attachment);

                                        // attachments can be either message/rfc822 parts or regular MIME parts
                                        if (entity is MimeKit.MessagePart)
                                        {
                                            var rfc822 = (MimeKit.MessagePart)entity;

                                            var path = Path.Combine(directory, attachment.PartSpecifier + ".eml");

                                            rfc822.Message.WriteTo(path);
                                        }
                                        else
                                        {
                                            var part = (MimePart)entity;

                                            // note: it's possible for this to be null, but most will specify a filename
                                            var fileName = part.FileName;

                                            var path = Path.Combine(directory, fileName);

                                            // decode and save the content to a file
                                            using (var stream = File.Create(path))
                                                part.Content.DecodeTo(stream);
                                        }
                                    }
                                    j++;
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

    }
}

