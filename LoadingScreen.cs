using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Windows.Forms;

namespace SSR___Case_Notes_Creator
{
    public partial class LoadingScreen : Form
    {
        private System.Windows.Forms.Timer timer;
        private int angle = 0; // Current rotation angle

        public LoadingScreen()
        {
            // Set up the form
            this.Text = "Loading...";
            this.Size = new Size(300, 300);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.White;
            this.DoubleBuffered = true; // Smooth animation

            // Initialize the timer
            timer = new System.Windows.Forms.Timer();
            timer.Interval = 30; // Lower value = smoother animation
            timer.Tick += (s, e) =>
            {
                angle += 10; // Increment rotation angle
                if (angle >= 360) angle = 0; // Reset angle after full rotation
                this.Invalidate(); // Trigger repaint
            };
            timer.Start();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            Graphics g = e.Graphics;

            // Enable anti-aliasing for smooth graphics
            g.SmoothingMode = SmoothingMode.AntiAlias;

            // Draw circular indicator
            int centerX = this.ClientSize.Width / 2;
            int centerY = this.ClientSize.Height / 2;
            int radius = 50; // Radius of the circle
            int lineThickness = 10; // Thickness of the lines

            using (Pen pen = new Pen(Color.DodgerBlue, lineThickness))
            {
                pen.StartCap = LineCap.Round;
                pen.EndCap = LineCap.Round;

                // Draw an arc for the loading effect
                g.DrawArc(pen, centerX - radius, centerY - radius, 2 * radius, 2 * radius, angle, 90);
            }
        }
    }
}
