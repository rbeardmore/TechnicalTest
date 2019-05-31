using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data.Entity;
using System.ComponentModel.DataAnnotations;


namespace TechnicalTestMVC.Models
{
    public class Customer
    {
        public Customer()
        {

        }
        public Customer(Guid guid, string fullName, DateTime dateCreated, DateTime dateOfBirth, double amount, int reference )
        {
            Guid = guid;
            fullName = FullName;
            DateCreated = dateCreated;
            DateOfBirth = dateOfBirth;
            Amount = amount;
            Ref = reference;
        }
        [Key]
        [Required(ErrorMessage = "Required Guid")]
        public Guid Guid { get; set; }
        [Required(ErrorMessage = "Required FullName")]
        [MaxLength(50)]
        public string FullName { get; set; }
        [Required(ErrorMessage = "Required DateCreated")]
        public DateTime DateCreated { get; set; }
        [Required(ErrorMessage = "Required DateOfBirth")]
        public DateTime DateOfBirth { get; set; }
        [Required(ErrorMessage = "Required Amount")]
        public double Amount { get; set; }
        [Required(ErrorMessage = "Required Ref")]
        public int Ref { get; set; }
    }
    public class CustomerDBContext : DbContext
    {
        public DbSet<Customer> Customers { get; set; }
    }
}