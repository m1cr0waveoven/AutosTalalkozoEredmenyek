using Newtonsoft.Json;

namespace AutosTalalkozoEredmenyek.Models;

internal sealed class Autolimbo
{
    [JsonProperty(PropertyName = "rendszam")]
    public string Rendszam { get; set; }

    [JsonProperty(PropertyName = "kategoria")]
    public string Kategoria { get; set; }

    [JsonProperty(PropertyName = "magassag")]
    public int Magassag { get; set; }
}
