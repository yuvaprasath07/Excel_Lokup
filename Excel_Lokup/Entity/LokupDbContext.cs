using Entity.Model;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Entity
{
    public class LokupDbContext : DbContext
    {
        public LokupDbContext(DbContextOptions<LokupDbContext> options) : base(options){}
        public DbSet<Area> L_Area { get; set; }
        public DbSet<Loadpattern> L_Loadpattern { get; set; }
        public DbSet<MainFeeStructure> L_MainFeeStructure {get; set; }
        public DbSet<PaymentMethod> L_PaymentMethod {get;set;}
        public DbSet<PowerSupplyContractTypesContracts2> L_PowerSupplyContractTypesContracts2 {get; set;}
        public DbSet<SelfSupplementaryPowerCalculationMethod> L_SelfSupplementaryPowerCalculationMethod { get; set; }
        public DbSet<Utility> L_Utility { get; set;}
    }
}
