using Newtonsoft.Json;

namespace AutosTalalkozoEredmenyek.Models;

internal sealed class Idomero
{
    [JsonProperty(PropertyName = "nev")]
    public string Nev { get; set; }

    [JsonProperty(PropertyName = "hibapont")]
    public int Hibapont { get; set; }

    [JsonProperty(PropertyName = "ido")]
    public int Ido { get; set; }
}
