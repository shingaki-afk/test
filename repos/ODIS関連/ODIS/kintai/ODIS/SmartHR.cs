using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Windows.Forms;

namespace ODIS.ODIS
{
    public partial class SmartHR : Form
    {

        private const string ApiKey = "shr_9139_1uJjoJyjogRU8xsuqyzQt1gq5srctbGY"; // 取得したAPIキーを設定
        private const string BaseUrl = "https://oki-daiken.smarthr.jp/v1/";


        public SmartHR()
        {
            InitializeComponent();
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", ApiKey);
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                try
                {
                    // ドロップダウンリスト項目のカスタム項目テンプレートのJSONデータ
                    var customItemTemplate = new
                    {
                        name = "現場名",
                        description = "現場名を選択してください。",
                        item_type = "enum", // ドロップダウンリスト項目
                        required = true,
                        options = new
                        {
                            choices = new object[]
                            {
                                new { label = "東京本社", value = "tokyo_hq" },
                                new { label = "大阪支店", value = "osaka_branch" },
                                new { label = "福岡営業所", value = "fukuoka_office" }
                            }
                        }
                    };

                    string json = JsonConvert.SerializeObject(customItemTemplate);
                    var content = new StringContent(json, Encoding.UTF8, "application/json");

                    HttpResponseMessage response;
                    if (string.IsNullOrEmpty(textBox2.Text))
                    {
                        // 登録
                        response = await client.PostAsync(BaseUrl + "employee_custom_item_templates", content);
                    }
                    else
                    {
                        // 更新
                        response = await client.PutAsync(BaseUrl + "employee_custom_item_templates/" + textBox2.Text, content);
                    }

                    response.EnsureSuccessStatusCode();

                    string responseJson = await response.Content.ReadAsStringAsync();
                    JObject result = JsonConvert.DeserializeObject<JObject>(responseJson);

                    // 結果をテキストボックスに表示
                    textBox1.Text = result.ToString();
                }
                catch (HttpRequestException ex)
                {
                    MessageBox.Show($"APIリクエストエラー: {ex.Message}");
                }
                catch (JsonException ex)
                {
                    MessageBox.Show($"JSONパースエラー: {ex.Message}");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"エラーが発生しました: {ex.Message}");
                }
            }
        }
    }
}
