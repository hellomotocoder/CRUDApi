using BCrypt;
using CRUDApi.Data;
using CRUDApi.Models;
using CRUDApi.Utils;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.IdentityModel.Tokens;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IdentityModel.Tokens.Jwt;
using System.Linq;
using System.Security.Claims;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using DinkToPdf;
using DinkToPdf.Contracts;
using System;
using System.IO;
using System.Threading.Tasks;

namespace CRUDApi.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class UsersController : ControllerBase
    {
        private readonly ApplicationDbContext _context;
        private readonly IConfiguration _configuration;

        public UsersController(ApplicationDbContext context, IConfiguration configuration)
        {
            _context = context;
            _configuration = configuration;
        }

        [HttpPost("login")]
        public IActionResult Login([FromBody] LoginModel loginModel)
        {
            var user = _context.Users.SingleOrDefault(u => u.Username == loginModel.Username);

            if (user == null || !BCrypt.Net.BCrypt.Verify(loginModel.Password, user.Password))
                return Unauthorized(new { StatusCode = StatusCodes.Status401Unauthorized, Message = "Invalid username or password" });

            var token = JwtUtils.GenerateJwtToken(user, _configuration["Jwt:Secret"]);

            return Ok(new { Token = token });
        }

        [HttpPost("register")]
        [AllowAnonymous]
        public async Task<IActionResult> RegisterUser([FromBody] User newUser)
        {
            try
            {
                // Validate the request body
                if (!ModelState.IsValid)
                {
                    return BadRequest(new { StatusCode = StatusCodes.Status400BadRequest, Message = "Invalid request body" });
                }

                // Ensure the username is unique
                if (_context.Users.Any(u => u.Username == newUser.Username))
                {
                    return BadRequest(new { StatusCode = StatusCodes.Status400BadRequest, Message = "Username is already taken" });
                }

                // Hash the password before saving to the database
                newUser.Password = BCrypt.Net.BCrypt.HashPassword(newUser.Password);

                // Add the new user to the database
                _context.Users.Add(newUser);
                await _context.SaveChangesAsync();

                return CreatedAtAction(nameof(GetUserById), new { userId = newUser.Id }, newUser);
            }
            catch (Exception ex)
            {
                // Log the exception
                return StatusCode(StatusCodes.Status500InternalServerError, new { StatusCode = StatusCodes.Status500InternalServerError, Message = "Internal Server Error" });
            }
        }

        [HttpGet("all")]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> GetAllUsers()
        {
            try
            {
                var users = await _context.Users.ToListAsync();
                return Ok(users);
            }
            catch (Exception ex)
            {
                // Log the exception
                return StatusCode(StatusCodes.Status500InternalServerError, new { StatusCode = StatusCodes.Status500InternalServerError, Message = "Internal Server Error" });
            }
        }

        // Add other CRUD operations (Create, Update, Delete) following a similar pattern.

        // Bonus Point 7: Implement Search Users endpoint
        [HttpPost("search")]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> SearchUsers([FromBody] SearchModel searchModel)
        {
            try
            {
                // Validate the request body
                if (!ModelState.IsValid)
                {
                    return BadRequest(new { StatusCode = StatusCodes.Status400BadRequest, Message = "Invalid request body" });
                }

                var query = _context.Users.AsQueryable();

                // Apply search filters
                if (!string.IsNullOrEmpty(searchModel.FieldName) && !string.IsNullOrEmpty(searchModel.FieldValue))
                {
                    switch (searchModel.FieldName.ToLower())
                    {
                        case "username":
                            query = query.Where(u => u.Username.Contains(searchModel.FieldValue));
                            break;

                        case "hobbies":
                            query = query.Where(u => u.Hobbies.Contains(searchModel.FieldValue));
                            break;

                        // Add other filter cases as needed

                        default:
                            return BadRequest(new { StatusCode = StatusCodes.Status400BadRequest, Message = "Invalid Field Name" });
                    }
                }

                // Apply pagination
                int pageSize = 10; // Set your desired page size
                int pageNumber = searchModel.PageNumber ?? 1; // Default to page 1 if not provided

                query = query.Skip((pageNumber - 1) * pageSize).Take(pageSize);

                // Apply sorting
                if (!string.IsNullOrEmpty(searchModel.SortBy))
                {
                    switch (searchModel.SortBy.ToLower())
                    {
                        case "username":
                            query = query.OrderBy(u => u.Username);
                            break;

                        case "age":
                            query = query.OrderBy(u => u.Age);
                            break;

                        // Add other sorting cases as needed

                        default:
                            return BadRequest(new { StatusCode = StatusCodes.Status400BadRequest, Message = "Invalid Sort By field" });
                    }
                }

                var filteredUsers = await query.ToListAsync();

                // Return the result
                return Ok(filteredUsers);
            }
            catch (Exception ex)
            {
                // Log the exception
                Console.WriteLine($"Error in SearchUsers: {ex.Message}");
                return StatusCode(StatusCodes.Status500InternalServerError, new { StatusCode = StatusCodes.Status500InternalServerError, Message = "Internal Server Error" });
            }
        }


        // Bonus Point 8: Implement Export Users endpoint
        [HttpPost("export")]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> ExportUsers([FromBody] SearchModel searchModel)
        {
            try
            {
                // Implement export logic to PDF or Excel based on search filters

                var filteredUsers = await _context.Users
                    .Where(u => u.Username.Contains(searchModel.FieldValue) || u.Hobbies.Contains(searchModel.FieldValue))
                    // Add other filtering criteria as needed
                    .ToListAsync();

                if (filteredUsers.Count == 0)
                {
                    return BadRequest(new { StatusCode = StatusCodes.Status400BadRequest, Message = "No users found for export" });
                }

                // Create Excel workbook
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Users");

                    // Add header: Current Date
                    worksheet.Cell(1, 1).Value = "Current Date:";
                    worksheet.Cell(1, 2).Value = DateTime.Now.ToShortDateString();

                    // Add header: Column Names
                    worksheet.Cell(2, 1).Value = "Username";
                    worksheet.Cell(2, 2).Value = "Age";
                    // Add other columns as needed

                    // Add data rows
                    for (int i = 0; i < filteredUsers.Count; i++)
                    {
                        worksheet.Cell(i + 3, 1).Value = filteredUsers[i].Username;
                        worksheet.Cell(i + 3, 2).Value = filteredUsers[i].Age;
                        // Add other columns as needed
                    }

                    // Add footer: Page Number manually
                    var pageSetup = worksheet.PageSetup;
                    pageSetup.Footer.Left.AddText("Page &P of &N").SetBold().SetItalic();

                    // Save the workbook to a MemoryStream
                    using (var excelStream = new MemoryStream())
                    {
                        workbook.SaveAs(excelStream);

                        // Convert Excel to PDF
                        var pdfDocument = new HtmlToPdfDocument
                        {
                            GlobalSettings = { PaperSize = PaperKind.A4 },
                            Objects =
                    {
                        new ObjectSettings
                        {
                            HtmlContent = " ",
                            WebSettings = { DefaultEncoding = "utf-8" }
                        }
                    }
                        };

                        var pdfBytes = new SynchronizedConverter(new PdfTools()).Convert(pdfDocument);

                        // Combine PDF and Excel
                        using (var combinedStream = new MemoryStream())
                        {
                            excelStream.Position = 0;
                            excelStream.CopyTo(combinedStream);

                            // Write PDF bytes to the combined stream
                            combinedStream.Write(pdfBytes, 0, pdfBytes.Length);

                            // Set the position to the beginning of the stream
                            combinedStream.Position = 0;

                            // Return the file
                            return File(combinedStream.ToArray(), "application/octet-stream", "exported_users.xlsx");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Log the exception
                Console.WriteLine($"Error in ExportUsers: {ex.Message}");
                return StatusCode(StatusCodes.Status500InternalServerError, new { StatusCode = StatusCodes.Status500InternalServerError, Message = "Internal Server Error" });
            }
        }

        [HttpGet("{userId}")]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> GetUserById(Guid userId)
        {
            try
            {
                var user = await _context.Users.FindAsync(userId);

                if (user == null)
                {
                    return NotFound(new { StatusCode = StatusCodes.Status404NotFound, Message = "User not found" });
                }

                return Ok(user);
            }
            catch (Exception ex)
            {
                // Log the exception
                return StatusCode(StatusCodes.Status500InternalServerError, new { StatusCode = StatusCodes.Status500InternalServerError, Message = "Internal Server Error" });
            }
        }


        [HttpPost("create")]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> CreateUser([FromBody] User newUser)
        {
            try
            {
                // Validate the request body
                if (!ModelState.IsValid)
                {
                    return BadRequest(new { StatusCode = StatusCodes.Status400BadRequest, Message = "Invalid request body" });
                }

                // Ensure the username is unique
                if (_context.Users.Any(u => u.Username == newUser.Username))
                {
                    return BadRequest(new { StatusCode = StatusCodes.Status400BadRequest, Message = "Username is already taken" });
                }

                // Hash the password before saving to the database
                newUser.Password = BCrypt.Net.BCrypt.HashPassword(newUser.Password);

                // Add the new user to the database
                _context.Users.Add(newUser);
                await _context.SaveChangesAsync();

                return CreatedAtAction(nameof(GetUserById), new { userId = newUser.Id }, newUser);
            }
            catch (Exception ex)
            {
                // Log the exception
                return StatusCode(StatusCodes.Status500InternalServerError, new { StatusCode = StatusCodes.Status500InternalServerError, Message = "Internal Server Error" });
            }
        }

        [HttpPut("update/{userId}")]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> UpdateUser(Guid userId, [FromBody] User updatedUser)
        {
            try
            {
                // Validate the request body
                if (!ModelState.IsValid)
                {
                    return BadRequest(new { StatusCode = StatusCodes.Status400BadRequest, Message = "Invalid request body" });
                }

                // Check if the user with the specified ID exists
                var existingUser = await _context.Users.FindAsync(userId);
                if (existingUser == null)
                {
                    return NotFound(new { StatusCode = StatusCodes.Status404NotFound, Message = "User not found" });
                }

                // Update properties of the existing user
                existingUser.Username = updatedUser.Username;
                existingUser.Password = BCrypt.Net.BCrypt.HashPassword(updatedUser.Password); // Ensure to hash the updated password
                existingUser.IsAdmin = updatedUser.IsAdmin;
                existingUser.Age = updatedUser.Age;
                existingUser.Hobbies = updatedUser.Hobbies;

                // Save changes to the database
                await _context.SaveChangesAsync();

                return Ok(existingUser);
            }
            catch (Exception ex)
            {
                // Log the exception
                return StatusCode(StatusCodes.Status500InternalServerError, new { StatusCode = StatusCodes.Status500InternalServerError, Message = "Internal Server Error" });
            }
        }

        [HttpDelete("delete/{userId}")]
        [Authorize(Roles = "Admin")]
        public async Task<IActionResult> DeleteUser(Guid userId)
        {
            try
            {
                // Check if the user with the specified ID exists
                var userToDelete = await _context.Users.FindAsync(userId);
                if (userToDelete == null)
                {
                    return NotFound(new { StatusCode = StatusCodes.Status404NotFound, Message = "User not found" });
                }

                // Remove the user from the database
                _context.Users.Remove(userToDelete);
                await _context.SaveChangesAsync();

                return NoContent();
            }
            catch (Exception ex)
            {
                // Log the exception
                return StatusCode(StatusCodes.Status500InternalServerError, new { StatusCode = StatusCodes.Status500InternalServerError, Message = "Internal Server Error" });
            }
        }
    }
}