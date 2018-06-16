using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace TestAuthCookies.Models
{
    public class PostConsentViewModel
    {
        /// <inheritdoc/>
        public PostConsentViewModel()
        {
        }

        /// <summary>
        /// URL to be redirected either the same window or closing the child and redirecting in the parent
        /// </summary>
        public string RedirectUrl { get; set; }
    }
}