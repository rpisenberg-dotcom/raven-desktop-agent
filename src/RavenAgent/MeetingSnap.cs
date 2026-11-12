using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using RavenAgent.Core;

namespace RavenAgent;

/// <summary>
/// Ctrl+Alt+S: a picture of the screen you're looking at (the one with the window in front), for the meeting going on.
/// Only taken when you press the keys. Scaled to 1600px wide as a JPEG so it uploads quickly.
/// </summary>
static class MeetingSnap
{
    public static MeetingSnapRecord? Take(string? clientId)
    {
        try
        {
            var (hwnd, _, title) = Native.ForegroundWindow();
            var screen = hwnd != IntPtr.Zero ? Screen.FromHandle(hwnd) : Screen.PrimaryScreen!;
            var b = screen.Bounds;
            using var full = new Bitmap(b.Width, b.Height);
            using (var g = Graphics.FromImage(full)) g.CopyFromScreen(b.Location, Point.Empty, b.Size);
            var scale = Math.Min(1.0, 1600.0 / b.Width);
            using var small = new Bitmap(Math.Max(1, (int)(b.Width * scale)), Math.Max(1, (int)(b.Height * scale)));
            using (var g = Graphics.FromImage(small))
            {
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.DrawImage(full, 0, 0, small.Width, small.Height);
            }
            var jpeg = ImageCodecInfo.GetImageEncoders().First(c => c.FormatID == ImageFormat.Jpeg.Guid);
            using var ms = new MemoryStream();
            using var p = new EncoderParameters(1);
            p.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 78L);
            small.Save(ms, jpeg, p);
            return new MeetingSnapRecord
            {
                AgentKey = Guid.NewGuid().ToString("N"),
                CapturedAt = Json.Iso(DateTime.UtcNow),
                ImageBase64 = Convert.ToBase64String(ms.ToArray()),
                WindowTitle = title.Length > 200 ? title[..200] : title,
                ClientId = clientId,
            };
        }
        catch (Exception ex) { Log.Error("Screenshot failed", ex); return null; }
    }
}
