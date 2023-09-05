using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Entity.Model
{
    [Table("L_MainFeeStructure", Schema = "YUVA")]
    public class MainFeeStructure
    {
        [Key]

        public string? Description { get; set; }

    }
}
