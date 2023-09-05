using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Entity.Model
{
    [Table("L_PowerSupplyContractTypesContracts2", Schema = "YUVA")]
    public class PowerSupplyContractTypesContracts2
    {
        [Key]

        public string? Description { get; set; }

    }
}
