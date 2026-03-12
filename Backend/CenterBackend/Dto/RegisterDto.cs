namespace CenterBackend.Dto
{
    public class RegisterDto
    {
        public string UserAccount { get; set; } = string.Empty;

        public string UserPassword { get; set; } = string.Empty;

        public string CheckPassword { get; set; } = string.Empty;
    }
}
