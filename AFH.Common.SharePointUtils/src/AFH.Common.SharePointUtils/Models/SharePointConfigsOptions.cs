using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AFH.Common.SharePointUtils.Models;

public sealed class SharePointListConfig
{
    public string SiteId { get; set; } = string.Empty;
    public string ListId { get; set; } = string.Empty;
    public string? ConfigListId { get; set; } = string.Empty;

 
}


