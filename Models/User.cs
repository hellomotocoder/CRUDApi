using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;


namespace CRUDApi.Models
{
    public class User
    {
        public Guid Id { get; set; }

        [Required]
        public string Username { get; set; }

        [Required]
        public string Password { get; set; }

        public bool IsAdmin { get; set; }

        [Required]
        public int Age { get; set; }

        public List<string> Hobbies { get; set; } = new List<string>();
    }
}
