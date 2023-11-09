### 0. 기본 화면

📍 화면 구성
1. Parking RESTFul(Test)
2. PreGate
3. InGate(Main(IN))
4. OutGate(Main(OUT))
5. Excel Page
6. Auto_IN, Auto_OUT(Multi Threading)

![화면 캡처 2023-10-31 232743](https://github.com/JUSEOUNGHYUN/JUSEOUNGHYUN/assets/80812790/23cb2005-5dba-4344-acd1-f3da6e74fe13)

### 1. Excel Open

![Excel open](https://github.com/JUSEOUNGHYUN/JUSEOUNGHYUN/assets/80812790/ccd1d8a3-0e75-4f04-9d8d-1777c317a318)

#### 1.1 ParseExelOpenData()
         private List<string> ParseExelOpenData(string strSheetName)
         {
             List<string> InList = new List<string>();
             if (m_strFileName != "")
             {
                 Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
                 Workbook workbook = application.Workbooks.Open(Filename: m_strFileName);
                 if (workbook != null)
                 {
                     Worksheet worksheet1 = workbook.Worksheets.get_Item(strSheetName);
                     application.Visible = false;
                     Range range = worksheet1.UsedRange;
                     String data = string.Empty;
                     string str = string.Empty;
         
                     m_nRowCnt = range.Rows.Count;
                     m_nColCnt = range.Columns.Count;
         
                     for (int i = 1; i <= range.Rows.Count; ++i)
                     {
                         for (int j = 1; j <= range.Columns.Count; ++j)
                         {
                             if (((range.Cells[i, j] as Range).Value2 == null))
                             {
                                 str = " ";
                             }
                             else
                             {
                                 str = ((range.Cells[i, j] as Range).Value2.ToString() + " ");
                             }
                             data += str;
                         }
                         InList = data.Split(' ').ToList();
                     }
                     InList.RemoveAt(InList.Count - 1);
         
                     workbook.Close(Filename: m_strFileName);
                     DeleteObject(workbook);
                     DeleteObject(worksheet1);
                     application.Quit();
                     DeleteObject(application);
                 }
                 else
                 {
                     MessageBox.Show("workbook이 null임.");
                 }
             }
             else
             {
                 MessageBox.Show("Failed to read file.");
             }
         
             return InList;
         }
### 2. Main InGate(RESTFul)

https://github.com/JUSEOUNGHYUN/JUSEOUNGHYUN/assets/80812790/4545ad25-6dc0-4789-915f-6de52e130fa7

#### 2.1 DoHttpWebRequest()
        private bool DoHttpWebRequest()
        {
            if (m_Request != null)
            {
                m_Request.Abort();
                m_Request = null;
            }
            m_URL = "http://" + IPADDRESS_textBox.Text + ":" + PORTNUMBER_textBox.Text + "/api/InOutCar";
            // 예외처리 IP와 PORT번호를 입력하지 않았을 때
            if (IPADDRESS_textBox.Text != string.Empty && PORTNUMBER_textBox.Text != string.Empty)
            {
                try
                {
                    // HTTP Websocket Communication
                    m_Request = (HttpWebRequest)WebRequest.Create(m_URL);
                    m_Request.Method = "POST"; // method 
                    m_Request.ContentType = "application/json"; // ContentType

                    m_IsIpPort = true;
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
            }
            else
            {
                m_IsIpPort = false;
            }
            return m_IsIpPort;
        }


#### 2.2 SendOrResultJson()
    private void SendOrResultJson(JObject obj)
    {
        if (DoHttpWebRequest())
        {
            m_StreamWriter = new StreamWriter(m_Request.GetRequestStream());
            m_StreamWriter.Write(obj);
            m_StreamWriter.Flush();
            m_StreamWriter.Close();
            SetFileWriteMsg("SendMessage");
            OutputJson(obj);
    
            try
            {
                HttpWebResponse httpResponse = (HttpWebResponse)m_Request.GetResponse();
                using (StreamReader streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    SetFileWriteMsg("ResponseMessage");
                    string result = streamReader.ReadToEnd();
                    ResultJson(result);
                    streamReader.Close();
                }
            }
            catch (Exception ex) // Error
            {
                Console.WriteLine(ex.Message);
            }
        }
    }

### 3. Auto IN (Multi Threading)
- Gate 101 102, 103(PreGate), 201, 202, 203(InGate) 301, 302, 303(OutGate) each Thread

https://github.com/JUSEOUNGHYUN/JUSEOUNGHYUN/assets/80812790/435249f9-0d7e-41be-b1ad-06fe5bea67dd

### 4. TextBox_TextChanged(UI)
- 컨테이너 사이즈가 40이 넘으면 안됨
- 첫 번째 문자는 N,L, 숫자만 가능

https://github.com/JUSEOUNGHYUN/JUSEOUNGHYUN/assets/80812790/6fc8260b-9ef2-4f33-b69e-b5d9f61b2aca

### 5. TabControl Design(DrawItem)

https://github.com/JUSEOUNGHYUN/JUSEOUNGHYUN/assets/80812790/4b0450d8-29cb-4cd4-b2d6-35319d5cecc0
