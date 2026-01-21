using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Windows.Forms;

namespace ODIS.ODIS
{
    public partial class SmartHR_Get : Form
    {

        private const string ApiKey = "shr_9139_1uJjoJyjogRU8xsuqyzQt1gq5srctbGY"; // 取得したAPIキーを設定
        private const string BaseUrl = "https://oki-daiken.smarthr.jp/api/v1/";


        public SmartHR_Get()
        {
            InitializeComponent();
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", ApiKey);

                try
                {

                    HttpResponseMessage response = await client.GetAsync(BaseUrl + "crew_custom_field_templates?page=1&per_page=10");
                    response.EnsureSuccessStatusCode();

                    string json = await response.Content.ReadAsStringAsync();
                    JArray result = JsonConvert.DeserializeObject<JArray>(json);

                    // 従業員カスタム項目テンプレート情報をテキストボックスに表示
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
