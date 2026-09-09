using System.ComponentModel.DataAnnotations;

namespace AkcniPlan.Models;

public enum TaskArea
{
    [Display(Name = "SVP")]
    Svp = 0,

    [Display(Name = "SDP")]
    Sdp = 1,

    [Display(Name = "BOZP")]
    Bozp = 2,

    [Display(Name = "PO")]
    Po = 3,

    [Display(Name = "Jiné")]
    Jine = 4
}