using Newtonsoft.Json;
using System;
using System.Collections.Generic;

namespace AutosTalalkozoEredmenyek;

internal interface IResultModel<T>
{
    string Error { get; }
    string Message { get; }
    IList<T> Results { get; }
}

internal sealed class ResultModel<T> : IResultModel<T>
{
    [JsonProperty(PropertyName = "error")]
    public string Error { get; set; }

    [JsonProperty(PropertyName = "message")]
    public string Message { get; set; }

    [JsonProperty(PropertyName = "results")]
    public IList<T> Results { get; set; }
}

internal sealed class NoResultModel<T> : IResultModel<T>
{
    [JsonProperty(PropertyName = "error")]
    public string Error { get; set; } = "Hiba az eredmény lekérdezése során";

    [JsonProperty(PropertyName = "message")]
    public string Message { get; set; } = "A lekérdezés üres eredménnyel tért vissza";

    [JsonProperty(PropertyName = "results")]
    public IList<T> Results => Array.Empty<T>();
}
