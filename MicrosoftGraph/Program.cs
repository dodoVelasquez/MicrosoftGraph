using Newtonsoft.Json;
using System.Net.Http.Headers;
using System.Text.Json;

/* Sustituir por los datos de aplicacion que tenga accesos */
string clientId = "TU_CLIENT_ID";
string clientSecret = "TU_CLIENT_SECRET";
string tenantId = "TU_TENANT_ID";

try
{
    /* Obtener token */
    var token = await GetTokenAsync();

    var httpClient = new HttpClient();

    /* Agregar el token de acceso a la solicitud */
    httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token.access_token);
    /* preferencia de respuesta en JSON */
    httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

    /* GET con filtro para optimizar la consulta */
    var responseGet = await httpClient.GetAsync(
        "https://graph.microsoft.com/v1.0/users?$select=Id,displayName,givenname,surname,employeeId,mail,jobTitle,officeLocation,department");

    /* Verificar si la solicitud fue exitosa */
    if (responseGet.IsSuccessStatusCode)
    {
        var responseBody = await responseGet.Content.ReadAsStringAsync();
        var jsonDoc = JsonDocument.Parse(responseBody);

        foreach (var usuario in jsonDoc.RootElement.GetProperty("value").EnumerateArray())
        {
            var id = usuario.GetProperty("id").GetString();
            var displayName = usuario.GetProperty("displayName").GetString();
            Console.WriteLine($"id: {id} Name: {displayName}");

            string? employeeIdUpdate = "0001";
            string? mailUpdate = "employee@contoso.com";
            string? departmentUpdate = "IT Department";

            /* Estructura para actualizar ficha office 365 */
            var userUpdate = new
            {
                employeeId = employeeIdUpdate,
                mail = mailUpdate,
                department = departmentUpdate
            };

            /* API con  patch para actualizar ficha office 365 */
            ////Comentado para evitar que se cometa algun error actualizando 
            // var patchUser = new HttpRequestMessage(HttpMethod.Patch, $"https://graph.microsoft.com/v1.0/users/{id}");
            // patchUser.Content = new StringContent(JsonConvert.SerializeObject(userUpdate), Encoding.UTF8, "application/json");
            // var responseUser = await httpClient.SendAsync(patchUser);

        }
        Console.WriteLine($"StatusCode: {Convert.ToInt32(responseGet.StatusCode)}  ReasonCode: {responseGet.ReasonPhrase}");
    }
    else
    {
        Console.WriteLine($"StatusCode: {Convert.ToInt32(responseGet.StatusCode)}  ReasonCode: {responseGet.ReasonPhrase}");
    }
}
catch (Exception ex)
{
    Console.WriteLine(ex.Message);
}

async Task<TokenResponse> GetTokenAsync()
{
    string url = "https://login.microsoftonline.com/" + tenantId + "/oauth2/v2.0/token";

    var values = new Dictionary<string, string>
    {
        { "client_id", clientId },
        { "scope", "https://graph.microsoft.com/.default" },
        { "client_secret", clientSecret },
        { "grant_type", "client_credentials" }
    };
    
    var data = new FormUrlEncodedContent(values);

    using var client = new HttpClient();
    var response = await client.PostAsync(url, data);
    string jsonToken = response.Content.ReadAsStringAsync().Result;

    TokenResponse result = JsonConvert.DeserializeObject<TokenResponse>(jsonToken);

    return result;
}

public class TokenResponse
{
    /* clase para obtener info token */
    public string access_token { get; set; }
    public string token_type { get; set; }
    public int expires_in { get; set; }
}