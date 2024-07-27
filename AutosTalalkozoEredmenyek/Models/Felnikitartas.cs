using Newtonsoft.Json;

namespace AutosTalalkozoEredmenyek.Models;

internal sealed class Felnikitartas
{
    [JsonProperty(PropertyName = "nev")]
    public string Nev { get; set; }

    [JsonProperty(PropertyName = "kategoria")]
    public string Versenynev { get; set; }

    [JsonProperty(PropertyName = "ido")]
    public int Ido { get; set; }
}
