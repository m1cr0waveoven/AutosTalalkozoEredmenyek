namespace AutosTalalkozoEredmenyek.Models;

internal sealed class Results
{
    public IResultModel<Autolimbo> Autolimbo { get; set; }
    public IResultModel<Felnikitartas> FelnikitartasNo { get; set; }
    public IResultModel<Felnikitartas> FelnikitartasFerfi { get; set; }
    public IResultModel<Idomero> Gumiguritas { get; set; }
    public IResultModel<Idomero> Autoszlalom { get; set; }
    public IResultModel<Autoszepsegverseny> Autoszepsegverseny { get; set; }
    public IResultModel<Idomero> Kviz { get; set; }
    public IResultModel<Idomero> AutoToloHuzo { get; set; }
    public IResultModel<Kipufogohangyomas> KipufogoHangnyomas { get; set; }
}
