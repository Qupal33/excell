$ErrorActionPreference = "Stop"

$source = @"
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;

public static class AppIconGenerator
{
    private sealed class IconFrame
    {
        public int Size { get; set; }
        public byte[] Bytes { get; set; }
    }

    public static void Generate(string buildDir)
    {
        Directory.CreateDirectory(buildDir);

        using (var preview = new Bitmap(1024, 1024, PixelFormat.Format32bppArgb))
        using (var graphics = Graphics.FromImage(preview))
        {
            DrawIcon(graphics, 1024);
            preview.Save(Path.Combine(buildDir, "icon.png"), ImageFormat.Png);
        }

        var sizes = new[] { 16, 24, 32, 48, 64, 128, 256 };
        var frames = new List<IconFrame>();

        foreach (var size in sizes)
        {
            frames.Add(CreateFrame(size));
        }

        SaveIco(frames, Path.Combine(buildDir, "icon.ico"));
    }

    private static IconFrame CreateFrame(int size)
    {
        using (var bitmap = new Bitmap(size, size, PixelFormat.Format32bppArgb))
        using (var graphics = Graphics.FromImage(bitmap))
        using (var stream = new MemoryStream())
        {
            DrawIcon(graphics, size);
            bitmap.Save(stream, ImageFormat.Png);
            return new IconFrame { Size = size, Bytes = stream.ToArray() };
        }
    }

    private static void SaveIco(List<IconFrame> frames, string outputPath)
    {
        using (var stream = new FileStream(outputPath, FileMode.Create, FileAccess.Write))
        using (var writer = new BinaryWriter(stream))
        {
            writer.Write((ushort)0);
            writer.Write((ushort)1);
            writer.Write((ushort)frames.Count);

            var offset = 6 + (16 * frames.Count);

            foreach (var frame in frames)
            {
                writer.Write((byte)(frame.Size >= 256 ? 0 : frame.Size));
                writer.Write((byte)(frame.Size >= 256 ? 0 : frame.Size));
                writer.Write((byte)0);
                writer.Write((byte)0);
                writer.Write((ushort)1);
                writer.Write((ushort)32);
                writer.Write(frame.Bytes.Length);
                writer.Write(offset);
                offset += frame.Bytes.Length;
            }

            foreach (var frame in frames)
            {
                writer.Write(frame.Bytes);
            }
        }
    }

    private static void DrawIcon(Graphics graphics, int size)
    {
        graphics.Clear(Color.Transparent);
        graphics.SmoothingMode = SmoothingMode.AntiAlias;
        graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
        graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
        graphics.CompositingQuality = CompositingQuality.HighQuality;

        var cardX = size * 0.11f;
        var cardY = size * 0.11f;
        var cardW = size * 0.78f;
        var cardH = size * 0.78f;
        var cardR = size * 0.18f;

        using (var shadowPath = RoundedRect(cardX + size * 0.03f, cardY + size * 0.035f, cardW, cardH, cardR))
        using (var shadowBrush = new SolidBrush(Color.FromArgb(58, 7, 18, 43)))
        {
            graphics.FillPath(shadowBrush, shadowPath);
        }

        using (var cardPath = RoundedRect(cardX, cardY, cardW, cardH, cardR))
        using (var bgBrush = new LinearGradientBrush(
            new PointF(cardX, cardY),
            new PointF(cardX + cardW, cardY + cardH),
            Hex("#0F766E"),
            Hex("#2563EB")))
        using (var borderPen = new Pen(Color.FromArgb(74, 255, 255, 255), Math.Max(2f, size * 0.012f)))
        {
            graphics.FillPath(bgBrush, cardPath);

            using (var overlayBrush = new PathGradientBrush(cardPath))
            {
                overlayBrush.CenterColor = Color.FromArgb(130, 255, 255, 255);
                overlayBrush.SurroundColors = new[] { Color.FromArgb(0, 255, 255, 255) };
                overlayBrush.CenterPoint = new PointF(cardX + cardW * 0.3f, cardY + cardH * 0.22f);
                graphics.FillPath(overlayBrush, cardPath);
            }

            graphics.DrawPath(borderPen, cardPath);
        }

        var docX = size * 0.29f;
        var docY = size * 0.20f;
        var docW = size * 0.42f;
        var docH = size * 0.56f;
        var docR = size * 0.045f;

        using (var docShadowPath = RoundedRect(docX + size * 0.012f, docY + size * 0.014f, docW, docH, docR))
        using (var docShadowBrush = new SolidBrush(Color.FromArgb(34, 6, 30, 52)))
        {
            graphics.FillPath(docShadowBrush, docShadowPath);
        }

        using (var docPath = RoundedRect(docX, docY, docW, docH, docR))
        using (var docBrush = new LinearGradientBrush(
            new PointF(docX, docY),
            new PointF(docX, docY + docH),
            Hex("#FDFEFE"),
            Hex("#E6F4FF")))
        using (var docPen = new Pen(Hex("#C6DCF5"), Math.Max(2f, size * 0.009f)))
        {
            graphics.FillPath(docBrush, docPath);
            graphics.DrawPath(docPen, docPath);
        }

        var fold = size * 0.115f;
        var foldPoints = new[]
        {
            new PointF(docX + docW - fold, docY),
            new PointF(docX + docW, docY),
            new PointF(docX + docW, docY + fold),
        };

        using (var foldBrush = new SolidBrush(Hex("#D5E7FF")))
        using (var foldPen = new Pen(Hex("#B5D0F0"), Math.Max(1f, size * 0.007f)))
        {
            graphics.FillPolygon(foldBrush, foldPoints);
            graphics.DrawLines(foldPen, new[]
            {
                new PointF(docX + docW - fold, docY),
                new PointF(docX + docW, docY),
                new PointF(docX + docW, docY + fold),
            });
        }

        var headerRect = new RectangleF(docX + docW * 0.11f, docY + docH * 0.11f, docW * 0.58f, docH * 0.08f);
        using (var headerPath = RoundedRect(headerRect.X, headerRect.Y, headerRect.Width, headerRect.Height, headerRect.Height * 0.45f))
        using (var headerBrush = new LinearGradientBrush(
            new PointF(headerRect.X, headerRect.Y),
            new PointF(headerRect.Right, headerRect.Y),
            Hex("#14B8A6"),
            Hex("#34D399")))
        {
            graphics.FillPath(headerBrush, headerPath);
        }

        using (var gridPen = new Pen(Color.FromArgb(110, 131, 166, 199), Math.Max(1f, size * 0.006f)))
        {
            var lineStart = docX + docW * 0.12f;
            var lineEnd = docX + docW * 0.82f;

            foreach (var offset in new[] { 0.31f, 0.41f, 0.51f, 0.61f })
            {
                var y = docY + docH * offset;
                graphics.DrawLine(gridPen, lineStart, y, lineEnd, y);
            }

            foreach (var offset in new[] { 0.30f, 0.48f, 0.66f })
            {
                var x = docX + docW * offset;
                graphics.DrawLine(gridPen, x, docY + docH * 0.30f, x, docY + docH * 0.66f);
            }
        }

        var barRects = new[]
        {
            new RectangleF(docX + docW * 0.14f, docY + docH * 0.58f, docW * 0.10f, docH * 0.13f),
            new RectangleF(docX + docW * 0.28f, docY + docH * 0.50f, docW * 0.10f, docH * 0.21f),
            new RectangleF(docX + docW * 0.42f, docY + docH * 0.41f, docW * 0.10f, docH * 0.30f),
        };

        var barColors = new[] { "#34D399", "#10B981", "#059669" };
        for (var i = 0; i < barRects.Length; i++)
        {
            using (var barPath = RoundedRect(barRects[i].X, barRects[i].Y, barRects[i].Width, barRects[i].Height, barRects[i].Width * 0.33f))
            using (var barBrush = new SolidBrush(Hex(barColors[i])))
            {
                graphics.FillPath(barBrush, barPath);
            }
        }

        var trendPoints = new[]
        {
            new PointF(docX + docW * 0.18f, docY + docH * 0.57f),
            new PointF(docX + docW * 0.34f, docY + docH * 0.47f),
            new PointF(docX + docW * 0.52f, docY + docH * 0.51f),
            new PointF(docX + docW * 0.74f, docY + docH * 0.30f),
        };

        using (var trendPen = new Pen(Hex("#F59E0B"), Math.Max(3f, size * 0.02f)))
        {
            trendPen.StartCap = LineCap.Round;
            trendPen.EndCap = LineCap.Round;
            trendPen.LineJoin = LineJoin.Round;
            graphics.DrawLines(trendPen, trendPoints);
        }

        using (var nodeBrush = new SolidBrush(Hex("#FBBF24")))
        using (var nodePen = new Pen(Color.FromArgb(180, 120, 53, 15), Math.Max(1f, size * 0.006f)))
        {
            var radius = size * 0.022f;
            foreach (var point in trendPoints)
            {
                graphics.FillEllipse(nodeBrush, point.X - radius, point.Y - radius, radius * 2f, radius * 2f);
                graphics.DrawEllipse(nodePen, point.X - radius, point.Y - radius, radius * 2f, radius * 2f);
            }
        }

        using (var sparkBrush = new SolidBrush(Color.FromArgb(210, 255, 243, 199)))
        {
            var sparkSize = size * 0.05f;
            var centerX = cardX + cardW * 0.78f;
            var centerY = cardY + cardH * 0.26f;
            graphics.FillEllipse(sparkBrush, centerX - sparkSize * 0.32f, centerY - sparkSize * 0.32f, sparkSize * 0.64f, sparkSize * 0.64f);
        }
    }

    private static GraphicsPath RoundedRect(float x, float y, float width, float height, float radius)
    {
        var diameter = Math.Max(1f, radius * 2f);
        var path = new GraphicsPath();

        path.AddArc(x, y, diameter, diameter, 180, 90);
        path.AddArc(x + width - diameter, y, diameter, diameter, 270, 90);
        path.AddArc(x + width - diameter, y + height - diameter, diameter, diameter, 0, 90);
        path.AddArc(x, y + height - diameter, diameter, diameter, 90, 90);
        path.CloseFigure();

        return path;
    }

    private static Color Hex(string hex, int alpha = 255)
    {
        hex = hex.TrimStart('#');
        var r = Convert.ToInt32(hex.Substring(0, 2), 16);
        var g = Convert.ToInt32(hex.Substring(2, 2), 16);
        var b = Convert.ToInt32(hex.Substring(4, 2), 16);
        return Color.FromArgb(alpha, r, g, b);
    }
}
"@

Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition $source

$buildDir = Split-Path -Parent $MyInvocation.MyCommand.Path
[AppIconGenerator]::Generate($buildDir)

Write-Output "Generated icon assets in $buildDir"
