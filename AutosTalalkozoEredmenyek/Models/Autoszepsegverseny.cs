using Newtonsoft.Json;

namespace AutosTalalkozoEredmenyek.Models;

internal sealed class Autoszepsegverseny
{
    [JsonProperty(PropertyName = "rendszam")]
    public string Rendszam { get; set; }

    [JsonProperty(PropertyName = "kulso")]
    public int Kulso { get; set; }

    [JsonProperty(PropertyName = "belso")]
    public int Belso { get; set; }

    [JsonProperty(PropertyName = "motorter")]
    public int Motorter { get; set; }

    [JsonProperty(PropertyName = "felni")]
    public int Felni { get; set; }

    [JsonProperty(PropertyName = "oszhang")]
    public int Osszhang { get; set; }

    [JsonProperty(PropertyName = "osszpontszam")]
    public int Osszpontszam { get; set; }
}
