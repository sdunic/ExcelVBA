using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Media;
using System.IO;
using System.Threading;
using System.Diagnostics;

namespace _2dGameLanguage
{
    public partial class BGL : Form
    {
        bool figureMoves;
        int offsetFieldFigure = 29;
        int boardOffsetX = 350, boardOffsetY = 760;
        int tempLineOffset = 5;
        int moves = 0;
        int numOfPlayers = 0;
        string[] players = {"Crveni", "Žuti", "Zeleni", "Plavi"};
        int activePlayer = 0; //vezano za listu igrača koji se prikazuju, 0 - Crveni ... 3 - Plavi
   	
    	
        //Instance Variables
        #region
        double lastTime, thisTime, diff;
        Sprite[] sprites = new Sprite[1000];
        Sprite[] redFigures = new Sprite[4];
        Sprite[] yellowFigures = new Sprite[4];
        Sprite[] greenFigures = new Sprite[4];
        Sprite[] blueFigures = new Sprite[4];
        Sprite[] redHomes = new Sprite[8];
        Sprite[] yellowHomes = new Sprite[8];
        Sprite[] greenHomes = new Sprite[8];
        Sprite[] blueHomes = new Sprite[8];
        Path[] pathCoordinates = new Path[40];  //koordinate po kojima ćemo se kretati
        SoundPlayer[] sounds = new SoundPlayer[1000];
        TextReader[] readFiles = new StreamReader[1000];
        TextWriter[] writeFiles = new StreamWriter[1000];
        int spriteCount = 0, soundCount = 0;
        string inkey;
        int mouseKey, mouseXp, mouseYp;
        Rectangle Collision;
        bool showSync = false;
        int loopcount;
        DateTime dt = new DateTime();
        String time;
        #endregion

        //Structs
        #region 
        public struct Sprite
        {
            public string image;
            public Bitmap bmp;
            public int x, y, width, height;
            public bool show;
            public bool moves;
            public int pathPosition;
            public int pathOffset;

            public Sprite(string images, int p1, int p2)
            {
                bmp = new Bitmap(images);
                image = images;
                x = p1;
                y = p2;
                width = bmp.Width;
                height = bmp.Height;
                show = true;
                moves = false;
                pathPosition = -1;
                pathOffset = -1;
            }

            public Sprite(string images, int p1, int p2, int w, int h)
            {
                bmp = new Bitmap(images);
                image = images;
                x = p1;
                y = p2;
                width = w;
                height = h;
                show = true;
                moves = false;
                pathPosition = -1;
                pathOffset = -1;
            }
        }

        public struct Path
        {
            public int x, y;
            public Path( int p1, int p2)
            {
                x = p1;
                y = p2;
            }
        }

        #endregion

        public BGL()
        {
            InitializeComponent();
        }

        public void Init()
        {
            while (numOfPlayers < 1 || numOfPlayers > 4)
            {
                numOfPlayers = ShowDialog("Broj igrača (1-4):", "Pitanje");
            }
            if (dt == null) time = dt.TimeOfDay.ToString();
            loopcount++;

            // definiramo moguće pomak čovečuljka na false, dakle stoji čovječuljak
            figureMoves = false;

                  
            setTitle("Čovječe ne ljuti se!!!"); // postavi naziv prozora
            setStatus("Baci kocku!");
            setBackgroundColour(209, 182, 137);
            loadSound(1, "roll-dice.wav"); // učitaj zvukove isto kao i sprite-ove

            SetBoard();
            SetHomes(numOfPlayers);
            loadSprite("direction.png", 91, spriteX(0), spriteY(0) - tempLineOffset);
            loadSprite("direction.png", 92, spriteX(20) + tempLineOffset, spriteY(20));
            rotateSprite(92, 90);
            loadSprite("direction.png", 93, spriteX(40), spriteY(40) + tempLineOffset);
            rotateSprite(93, 180);
            loadSprite("direction.png", 94, spriteX(60) - tempLineOffset, spriteY(60));
            rotateSprite(94, 270);
            setPlayer(activePlayer);
            
        }

        public void SetBoard() {

            int tempOffsetX = 0;
            int tempOffsetY = 0;
            
            //60px ima svaki field - učitaj sprite sa identifikatorom na određenu koordinatu

            for (int i = 0; i < 4; i++)
            {

                if (i == 0) 
                { 
                    tempOffsetX = boardOffsetX;
                    tempOffsetY = boardOffsetY;
                }
                if (i == 1)
                {
                    tempOffsetX = boardOffsetX - 315;
                    tempOffsetY = boardOffsetY - 441;
                }
                else if (i == 2)
                {
                    tempOffsetX = boardOffsetX + 127;
                    tempOffsetY = boardOffsetY - 756;
                }
                else if (i == 3)
                {
                    tempOffsetX = boardOffsetX + 442;
                    tempOffsetY = boardOffsetY - 315;
                }

                SetPath(tempOffsetX, tempOffsetY, offsetFieldFigure, "WhiteField.png", "StartField.png", "v_line.png", "h_line.png", i);
            }
            
            //Postavljanje kockica
            SetDice(6);
        }

        public int RollDice(int x)
        {
            if (s.ElapsedMilliseconds > 2000)
            {
                s.Stop();
            }
            else if (s.ElapsedMilliseconds % 500 == 0)
            {
                Random rnd = new Random();
                x = rnd.Next(1, 7);
                SetDice(x);
            }
            return x;
        }

        public void SetDice(int x)
        {
            loadSprite("dice-" + x.ToString() + ".png", 81, 875, 650);
        }

        public void SetPath(int offsetX, int offsetY, int offsetFieldFigure, string whitePath, string startPath, string verticalLine, string horizontalLine, int rotation)
        {
            int tempX = offsetX;
            int tempY = offsetY;
            int tempOffset = 2 * offsetFieldFigure + tempLineOffset;
     

            string spriteName = "";

            for (int i = (0 + rotation * 20); i < (20 + rotation * 20); i++)
            {
                if (i % 20 == 0) spriteName = startPath;
                else spriteName = whitePath;


                if ((i >= (0 + rotation * 20) && i < (10 + rotation * 20)) || (i >= (18 + rotation * 20) && i < (20 + rotation * 20)))
                {
                    if (i == (9 + rotation * 20))
                    {
                        
                        spriteName = horizontalLine;
                        tempX -= tempLineOffset;
                        tempY += offsetFieldFigure;
                        if (rotation == 1)
                        {
                            spriteName = verticalLine;
                            tempX += (offsetFieldFigure + tempLineOffset);
                            tempY -= (offsetFieldFigure + tempLineOffset);
                        }
                        else if (rotation == 2)
                        {
                            tempX += (offsetFieldFigure * 2 + tempLineOffset);
                        }
                        else if (rotation == 3)
                        {
                            spriteName = verticalLine;
                            tempX += (offsetFieldFigure + tempLineOffset);
                            tempY += offsetFieldFigure;
                        }

                    }
                    else if (i % 2 == 1 & i != (18 + rotation * 20))
                    {
                        spriteName = verticalLine;
                        tempX += offsetFieldFigure;
                        tempY -= tempLineOffset;
                        if (rotation == 1)
                        {
                            spriteName = horizontalLine;
                            tempX += offsetFieldFigure;
                            tempY += offsetFieldFigure + tempLineOffset;
                        }
                        else if (rotation == 2)
                        {
                            tempY += (offsetFieldFigure*2 + tempLineOffset);
                        }
                        else if (rotation == 3)
                        {
                            spriteName = horizontalLine;
                            tempX -= (offsetFieldFigure + tempLineOffset);
                            tempY += (offsetFieldFigure + tempLineOffset);
                        }
                    }
                    else 
                    {
                        if (rotation == 0)
                            tempY -= tempOffset;
                        else if (rotation == 1)
                            tempX += tempOffset;
                        else if (rotation == 2)
                            tempY += tempOffset;
                        else
                            tempX -= tempOffset;
                    }
                }
                else if (i >= (10 + rotation * 20) && i < (20 + rotation * 20))
                {
                    if (i == (17 + rotation * 20))
                    {
                        spriteName = verticalLine;
                        tempX += offsetFieldFigure;
                        tempY -= tempLineOffset;
                        if (rotation == 1)
                        {
                            spriteName = horizontalLine;
                            tempX += offsetFieldFigure;
                            tempY += offsetFieldFigure + tempLineOffset;
                        }
                        else if (rotation == 2)
                        {
                            tempY += (offsetFieldFigure * 2 + tempLineOffset);
                        }
                        else if (rotation == 3)
                        {
                            spriteName = horizontalLine;
                            tempX -= (offsetFieldFigure + tempLineOffset);
                            tempY += (offsetFieldFigure + tempLineOffset);
                        }
                    }
                    else if (i % 2 == 1)
                    {
                        spriteName = horizontalLine;
                        tempX -= tempLineOffset;
                        tempY += offsetFieldFigure;
                        if (rotation == 1)
                        {
                            spriteName = verticalLine;
                            tempX += offsetFieldFigure + tempLineOffset;
                            tempY -= (offsetFieldFigure + tempLineOffset);
                        }
                        else if (rotation == 2)
                        {
                            tempX += (2 * offsetFieldFigure + tempLineOffset);
                        }
                        else if (rotation == 3) 
                        {
                            spriteName = verticalLine;
                            tempX += offsetFieldFigure + tempLineOffset;
                            tempY += offsetFieldFigure;
                        }
                    }
                    else
                    {
                        if (rotation == 0)
                            tempX -= tempOffset;
                        else if (rotation == 1)
                            tempY -= tempOffset;
                        else if (rotation == 2)
                            tempX += tempOffset;
                        else
                            tempY += tempOffset;
                    }
                }

                //crtanje puta i upisivanje korditana u listu
                loadSprite(spriteName, i, tempX - offsetFieldFigure, tempY - offsetFieldFigure);
                if (i % 2 == 0)
                {
                    
                    loadPathCoordinates(i/2, tempX - offsetFieldFigure, tempY - offsetFieldFigure);
                }

                //ispravke offseta nazad u crtanju linija između polja
                if ((i >= (0 + rotation * 20) && i < (10 + rotation * 20)) || (i >= (18 + rotation * 20) && i < (20 + rotation * 20)))
                {
                    if (i == (9 + rotation * 20))
                    {
                        tempX += tempLineOffset;
                        tempY -= offsetFieldFigure;
                        if (rotation == 1)
                        {
                            tempX -= (offsetFieldFigure + tempLineOffset);
                            tempY += (offsetFieldFigure + tempLineOffset);
                        }
                        else if (rotation == 2)
                        {
                            tempX -= (offsetFieldFigure * 2 + tempLineOffset);
                        }
                        else if (rotation == 3)
                        {
                            tempX -= (offsetFieldFigure + tempLineOffset);
                            tempY -= offsetFieldFigure;
                        }
                    }
                    else if (i % 2 == 1 & i != (18 + rotation * 20))
                    {
                        tempX -= offsetFieldFigure;
                        tempY += tempLineOffset;
                        if (rotation == 1)
                        {
                            tempX -= offsetFieldFigure;
                            tempY -= (offsetFieldFigure + tempLineOffset);
                        } 
                        else if (rotation == 2)
                        {
                            tempY -= (offsetFieldFigure * 2 + tempLineOffset);
                        }
                        else if (rotation == 3)
                        {
                            tempX += (offsetFieldFigure + tempLineOffset);
                            tempY -= (offsetFieldFigure + tempLineOffset);
                        }
                    }
                }
                else if (i >= (10 + rotation * 20) && i < (20 + rotation * 20))
                {
                    if (i == (17 + rotation * 20))
                    {
                        tempX -= offsetFieldFigure;
                        tempY += tempLineOffset;
                        if (rotation == 1)
                        {
                            tempX -= offsetFieldFigure;
                            tempY -= (offsetFieldFigure + tempLineOffset);
                        }
                        else if (rotation == 2)
                        {
                            tempY -= (offsetFieldFigure * 2 + tempLineOffset);
                        }
                        else if (rotation == 3)
                        {
                            tempX += (offsetFieldFigure + tempLineOffset);
                            tempY -= (offsetFieldFigure + tempLineOffset);
                        }
                    }
                    else if (i % 2 == 1)
                    {
                        tempX += tempLineOffset;
                        tempY -= offsetFieldFigure;
                        if (rotation == 1)
                        {
                            tempX -= offsetFieldFigure + tempLineOffset;
                            tempY += (offsetFieldFigure + tempLineOffset);   
                        }
                        else if (rotation == 2)
                        {
                            tempX -= (2 * offsetFieldFigure + tempLineOffset);
                        }
                        else if (rotation == 3)
                        {
                            tempX -= offsetFieldFigure + tempLineOffset;
                            tempY -= offsetFieldFigure;  
                        }
                    }
                }
               
            }
        }

        public void SetHomes(int numOfPlayers)
        {
            
            loadRedHomeSprite("RedField.png", 2, spriteX(14), spriteY(0));
            loadRedHomeSprite("RedField.png", 3, spriteX(16), spriteY(0));
            loadRedHomeSprite("RedField.png", 0, spriteX(14), spriteY(2));
            loadRedHomeSprite("RedField.png", 1, spriteX(16), spriteY(2));
            loadRedHomeSprite("RedField.png", 4, spriteX(78), spriteY(2));
            loadRedHomeSprite("RedField.png", 5, spriteX(78), spriteY(4));
            loadRedHomeSprite("RedField.png", 6, spriteX(78), spriteY(6));
            loadRedHomeSprite("RedField.png", 7, spriteX(78), spriteY(8));
            if (numOfPlayers > 0 && numOfPlayers < 5)
            {
                loadRedFigureSprite("RedFigure.png", 2, spriteX(14), spriteY(0));
                loadRedFigureSprite("RedFigure.png", 3, spriteX(16), spriteY(0));
                loadRedFigureSprite("RedFigure.png", 0, spriteX(14), spriteY(2));
                loadRedFigureSprite("RedFigure.png", 1, spriteX(16), spriteY(2));
            }

            loadYellowHomeSprite("YellowField.png", 1, spriteX(20), spriteY(34));
            loadYellowHomeSprite("YellowField.png", 0, spriteX(22), spriteY(34));
            loadYellowHomeSprite("YellowField.png", 3, spriteX(20), spriteY(36));
            loadYellowHomeSprite("YellowField.png", 2, spriteX(22), spriteY(36));
            loadYellowHomeSprite("YellowField.png", 4, spriteX(14), spriteY(18));
            loadYellowHomeSprite("YellowField.png", 5, spriteX(12), spriteY(18));
            loadYellowHomeSprite("YellowField.png", 6, spriteX(10), spriteY(18));
            loadYellowHomeSprite("YellowField.png", 7, spriteX(8), spriteY(18));
            if (numOfPlayers > 1 && numOfPlayers < 5)
            {
                loadYellowFigureSprite("YellowFigure.png", 1, spriteX(20), spriteY(34));
                loadYellowFigureSprite("YellowFigure.png", 0, spriteX(22), spriteY(34));
                loadYellowFigureSprite("YellowFigure.png", 3, spriteX(20), spriteY(36));
                loadYellowFigureSprite("YellowFigure.png", 2, spriteX(22), spriteY(36));
            }

            loadGreenHomeSprite("GreenField.png", 0, spriteX(54), spriteY(34));
            loadGreenHomeSprite("GreenField.png", 1, spriteX(56), spriteY(34));
            loadGreenHomeSprite("GreenField.png", 2, spriteX(54), spriteY(36));
            loadGreenHomeSprite("GreenField.png", 3, spriteX(56), spriteY(36));
            loadGreenHomeSprite("GreenField.png", 4, spriteX(78), spriteY(34));
            loadGreenHomeSprite("GreenField.png", 5, spriteX(78), spriteY(32));
            loadGreenHomeSprite("GreenField.png", 6, spriteX(78), spriteY(30));
            loadGreenHomeSprite("GreenField.png", 7, spriteX(78), spriteY(28));
            if (numOfPlayers > 2 && numOfPlayers < 5)
            {
                loadGreenFigureSprite("GreenFigure.png", 0, spriteX(54), spriteY(34));
                loadGreenFigureSprite("GreenFigure.png", 1, spriteX(56), spriteY(34));
                loadGreenFigureSprite("GreenFigure.png", 2, spriteX(54), spriteY(36));
                loadGreenFigureSprite("GreenFigure.png", 3, spriteX(56), spriteY(36));
            }

            loadBlueHomeSprite("BlueField.png", 2, spriteX(54), spriteY(0));
            loadBlueHomeSprite("BlueField.png", 3, spriteX(56), spriteY(0));
            loadBlueHomeSprite("BlueField.png", 0, spriteX(54), spriteY(2));
            loadBlueHomeSprite("BlueField.png", 1, spriteX(56), spriteY(2));
            loadBlueHomeSprite("BlueField.png", 4, spriteX(54), spriteY(18));
            loadBlueHomeSprite("BlueField.png", 5, spriteX(52), spriteY(18));
            loadBlueHomeSprite("BlueField.png", 6, spriteX(50), spriteY(18));
            loadBlueHomeSprite("BlueField.png", 7, spriteX(48), spriteY(18));
            if (numOfPlayers == 4)
            {
                loadBlueFigureSprite("BlueFigure.png", 2, spriteX(54), spriteY(0));
                loadBlueFigureSprite("BlueFigure.png", 3, spriteX(56), spriteY(0));
                loadBlueFigureSprite("BlueFigure.png", 0, spriteX(54), spriteY(2));
                loadBlueFigureSprite("BlueFigure.png", 1, spriteX(56), spriteY(2));
            }

        }

        private Stopwatch s = new Stopwatch();

      
        private void Update(object sender, EventArgs e)
        {        	
            // ponašanje figura 
            //startna pozicija sprite[0] crveni, sprite[10] zuti, sprite [20] zeleni, sprite [30] plavi
            
            if (isKeyDown(Keys.Enter) && !figureMoves)
            {
                s = new Stopwatch();
                s.Start();
                playSound(1);
                setStatus("Čekaj red!");
            }
            if (s.IsRunning)
            {
                moves = RollDice(moves);
            }
            else if(moves > 0)
            {
                if (activePlayer == 0)
                {
                    moveFigureToSprite(redFigures, 0, moves);
                }
                else if (activePlayer == 1)
                {
                    moveFigureToSprite(yellowFigures, 0, moves);
                }
                else if (activePlayer == 2)
                {
                    moveFigureToSprite(greenFigures, 0, moves);
                }
                else if (activePlayer == 3)
                {
                    moveFigureToSprite(blueFigures, 0, moves);
                }
                moves = 0;
                setStatus("Baci kocku!");
                
                activePlayer = (activePlayer + 1) % numOfPlayers;
                setPlayer(activePlayer);
            }

              	
            this.Refresh();
        }

        // Start of Game Methods

        #region

        //This is the beginning of the setter methods

        private void startTimer(object sender, EventArgs e)
        {
            timer1.Start();
            timer2.Start();
            Init();
        }

        public void showSyncRate(bool val)
        {
            showSync = val;
            if (val == true) syncRate.Show();
            if (val == false) syncRate.Hide();
        }


        public void updateSyncRate()
        {
            if (showSync == true)
            {
                thisTime = (DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0)).TotalSeconds;
                diff = thisTime - lastTime;
                lastTime = (DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0)).TotalSeconds;

                double fr = (1000 / diff) / 1000;

                int fr2 = Convert.ToInt32(fr);

                syncRate.Text = fr2.ToString();
            }
              
        }

        public void setTitle(string title)
        {
            this.Text = title;
        }

        public void setStatus(string status)
        {
            lblStatus.Text = status;
        }

        public void setPlayer(int pl)
        {
            lblPlayerInfo.Text = "Igrač - " + players[pl];
        }



        public void setBackgroundColour(int r, int g, int b)
        {
        
            this.BackColor = Color.FromArgb(r, g, b);
        }

        public void setBackgroundColour(Color col)
        {
            this.BackColor = col;
        }

        public void setBackgroundImage(string backgroundImage)
        {
            this.BackgroundImage = new Bitmap(backgroundImage);
        }

        public void setBackgroundImageLayout(string layout)
        {
            if (layout.ToLower() == "none") this.BackgroundImageLayout = ImageLayout.None;
            if (layout.ToLower() == "tile") this.BackgroundImageLayout = ImageLayout.Tile;
            if (layout.ToLower() == "stretch") this.BackgroundImageLayout = ImageLayout.Stretch;
            if (layout.ToLower() == "center") this.BackgroundImageLayout = ImageLayout.Center;
            if (layout.ToLower() == "zoom") this.BackgroundImageLayout = ImageLayout.Zoom;
        }
        
        private void updateFrameRate(object sender, EventArgs e)
        {
            updateSyncRate();
        }

        public void loadSprite(string file, int spriteNum)
        {
            spriteCount++;
            sprites[spriteNum] = new Sprite(file, 0, 0);
        }

        public void loadSprite(string file, int spriteNum, int x, int y)
        {
            spriteCount++;
            sprites[spriteNum] = new Sprite(file, x, y);
        }
        public void loadRedFigureSprite(string file, int spriteNum, int x, int y)
        {
            redFigures[spriteNum] = new Sprite(file, x, y);
            redFigures[spriteNum].pathOffset = 0;
        }
        public void loadYellowFigureSprite(string file, int spriteNum, int x, int y)
        {
            yellowFigures[spriteNum] = new Sprite(file, x, y);
            yellowFigures[spriteNum].pathOffset = 20;
        }
        public void loadGreenFigureSprite(string file, int spriteNum, int x, int y)
        {
            greenFigures[spriteNum] = new Sprite(file, x, y);
            greenFigures[spriteNum].pathOffset = 40;
        }
        public void loadBlueFigureSprite(string file, int spriteNum, int x, int y)
        {
            blueFigures[spriteNum] = new Sprite(file, x, y);
            blueFigures[spriteNum].pathOffset = 60;
        }
        public void loadRedHomeSprite(string file, int spriteNum, int x, int y)
        {
            redHomes[spriteNum] = new Sprite(file, x, y);
        }
        public void loadYellowHomeSprite(string file, int spriteNum, int x, int y)
        {
            yellowHomes[spriteNum] = new Sprite(file, x, y);
        }
        public void loadGreenHomeSprite(string file, int spriteNum, int x, int y)
        {
            greenHomes[spriteNum] = new Sprite(file, x, y);
        }
        public void loadBlueHomeSprite(string file, int spriteNum, int x, int y)
        {
            blueHomes[spriteNum] = new Sprite(file, x, y);
        }

        public void loadPathCoordinates(int pathNum, int x, int y)
        {
            pathCoordinates[pathNum] = new Path(x, y);
        }

        public void loadSprite(string file, int spriteNum, int x, int y, int w, int h)
        {
            spriteCount++;
            sprites[spriteNum] = new Sprite(file, x, y, w, h);
        }

        public void rotateSprite(int spriteNum, int angle)
        {
            if (angle == 90)
                sprites[spriteNum].bmp.RotateFlip(RotateFlipType.Rotate90FlipNone);
            if (angle == 180)
                sprites[spriteNum].bmp.RotateFlip(RotateFlipType.Rotate180FlipNone);
            if (angle == 270)
                sprites[spriteNum].bmp.RotateFlip(RotateFlipType.Rotate270FlipNone);
        }

        public void scaleSprite(int spriteNum, int scale)
        {
            float sx = float.Parse(sprites[spriteNum].width.ToString());
            float sy = float.Parse(sprites[spriteNum].height.ToString());
            float nsx = ((sx / 100) * scale); 
            float nsy = ((sy / 100) * scale);

            sprites[spriteNum].width = Convert.ToInt32(nsx);
            sprites[spriteNum].height = Convert.ToInt32(nsy);
        }

        public void moveSprite(Sprite[] tempSprites, int spriteNum, int x, int y)
        {
            tempSprites[spriteNum].x = x;
            tempSprites[spriteNum].y = y;
        }

        public void moveFigureToSprite(Sprite[] figures, int numOfFigure, int numOfMoves)
        {

            //ako se pomiče iz kućice postaviti na offset
            if (!figures[numOfFigure].moves && numOfMoves == 6)
            {
                figures[numOfFigure].x = sprites[figures[numOfFigure].pathOffset].x;
                figures[numOfFigure].y = sprites[figures[numOfFigure].pathOffset].y;
                figures[numOfFigure].moves = true;
                figures[numOfFigure].pathPosition = figures[numOfFigure].pathOffset;
            }
            else if(figures[numOfFigure].moves)
            {
                int newPosition = figures[numOfFigure].pathPosition + numOfMoves * 2;

                if (newPosition - figures[numOfFigure].pathOffset > 80)
                {
                    //ulazi u kućicu
                }
                else
                {
                    figures[numOfFigure].x = sprites[newPosition % 80].x;
                    figures[numOfFigure].y = sprites[newPosition % 80].y;
                    figures[numOfFigure].pathPosition = newPosition % 80;
                }
                
            }
        }

        public void setImageColorKey(int spriteNum, int r, int g, int b)
        {
            sprites[spriteNum].bmp.MakeTransparent(Color.FromArgb(r, g, b));
        }

        public void setImageColorKey(int spriteNum, Color col)
        {
            sprites[spriteNum].bmp.MakeTransparent(col);
        }

        public void setSpriteVisible(int spriteNum, bool ans)
        {
            sprites[spriteCount].show = ans;
        }

        public void hideSprite(int spriteNum)
        {
            sprites[spriteCount].show = false;
        }


        public void flipSprite(int spriteNum, string fliptype)
        {
            if(fliptype.ToLower() == "none")
            sprites[spriteNum].bmp.RotateFlip(RotateFlipType.RotateNoneFlipNone);

            if (fliptype.ToLower() == "horizontal")
            sprites[spriteNum].bmp.RotateFlip(RotateFlipType.RotateNoneFlipX);

            if (fliptype.ToLower() == "horizontalvertical")
            sprites[spriteNum].bmp.RotateFlip(RotateFlipType.RotateNoneFlipXY);

            if (fliptype.ToLower() == "vertical")
            sprites[spriteNum].bmp.RotateFlip(RotateFlipType.RotateNoneFlipY);
        }

        public void changeSpriteImage(int spriteNum, string file)
        {
            sprites[spriteNum] = new Sprite(file, sprites[spriteNum].x, sprites[spriteNum].y);
        }

        public void loadSound(int soundNum, string file)
        {
            soundCount++;
            sounds[soundNum] = new SoundPlayer(file);
        }

        public void playSound(int soundNum)
        {
            sounds[soundNum].Play();
        }

        public void loopSound(int soundNum)
        {
            sounds[soundNum].PlayLooping();
        }

        public void stopSound(int soundNum)
        {
            sounds[soundNum].Stop();
        }

        public void openFileToRead(string fileName, int fileNum)
        {
            readFiles[fileNum] = new StreamReader(fileName);
        }

        public void closeFileToRead(int fileNum)
        {
            readFiles[fileNum].Close();
        }

        public void openFileToWrite(string fileName, int fileNum)
        {
            writeFiles[fileNum] = new StreamWriter(fileName);
        }

        public void closeFileToWrite(int fileNum)
        {
            writeFiles[fileNum].Close();
        }

        public void writeLine(int fileNum, string line)
        {
            writeFiles[fileNum].WriteLine(line);
        }

        public void hideMouse()
        {
            Cursor.Hide();
        }

        public void showMouse()
        {
            Cursor.Show();
        }



        //This is the beginning of the getter methods

        public bool spriteExist(int spriteNum)
        {
            if (sprites[spriteNum].bmp != null)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public int spriteX(int spriteNum)
        {
            return sprites[spriteNum].x;
        }

        public int spriteY(int spriteNum)
        {
            return sprites[spriteNum].y;
        }

        public int spriteWidth(int spriteNum)
        {
            return sprites[spriteNum].width;
        }

        public int spriteHeight(int spriteNum)
        {
            return sprites[spriteNum].height;
        }

        public bool spriteVisible(int spriteNum)
        {
            return sprites[spriteNum].show;
        }

        public string spriteImage(int spriteNum)
        {
            return sprites[spriteNum].bmp.ToString();
        }

        public bool isKeyPressed(string key)
        {
            if (inkey == key)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public bool isKeyPressed(Keys key)
        {
            if (inkey == key.ToString())
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public bool isKeyDown(Keys key)
        {
            if (inkey == key.ToString())
            {
                inkey = "";
                return true;
            }
            else
            {
                return false;
            }
        }

        public bool spriteCollision(int spriteNum1, int spriteNum2)
        {
            Rectangle sp1 = new Rectangle(sprites[spriteNum1].x, sprites[spriteNum1].y, sprites[spriteNum1].width, sprites[spriteNum1].height);
            Rectangle sp2 = new Rectangle(sprites[spriteNum2].x, sprites[spriteNum2].y, sprites[spriteNum2].width, sprites[spriteNum2].height);
            Collision = Rectangle.Intersect(sp1, sp2);

            if (!Collision.IsEmpty)
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        public string readLine(int fileNum)
        {
            return readFiles[fileNum].ReadLine();
        }

        public string readFile(int fileNum)
        {
            return readFiles[fileNum].ReadToEnd();
        }

        public bool isMousePressed() {
            if (mouseKey == 1) return true;
            else return false;
        }

        public int mouseX()
        {
            return mouseXp;
        }

        public int mouseY()
        {
            return mouseYp;
        }

        #endregion


        //Game Update and Input
        #region
        private void Draw(object sender, PaintEventArgs e)
        {
            Graphics g = e.Graphics;

            foreach (Sprite sprite in sprites)
            {
                if (sprite.bmp != null && sprite.show == true)
                    g.DrawImage(sprite.bmp, new Rectangle(sprite.x, sprite.y, sprite.width, sprite.height));
            }
            //crtanje kuća ispred figurica tako da budu ispod
            foreach (Sprite sprite in redHomes)
            {
                if (sprite.bmp != null && sprite.show == true)
                    g.DrawImage(sprite.bmp, new Rectangle(sprite.x, sprite.y, sprite.width, sprite.height));
            }
            foreach (Sprite sprite in greenHomes)
            {
                if (sprite.bmp != null && sprite.show == true)
                    g.DrawImage(sprite.bmp, new Rectangle(sprite.x, sprite.y, sprite.width, sprite.height));
            }
            foreach (Sprite sprite in yellowHomes)
            {
                if (sprite.bmp != null && sprite.show == true)
                    g.DrawImage(sprite.bmp, new Rectangle(sprite.x, sprite.y, sprite.width, sprite.height));
            }
            foreach (Sprite sprite in blueHomes)
            {
                if (sprite.bmp != null && sprite.show == true)
                    g.DrawImage(sprite.bmp, new Rectangle(sprite.x, sprite.y, sprite.width, sprite.height));
            }
            //crtanje figurica
            foreach (Sprite sprite in redFigures)
            {
                if (sprite.bmp != null && sprite.show == true)
                    g.DrawImage(sprite.bmp, new Rectangle(sprite.x, sprite.y, sprite.width, sprite.height));
            }
            foreach (Sprite sprite in greenFigures)
            {
                if (sprite.bmp != null && sprite.show == true)
                    g.DrawImage(sprite.bmp, new Rectangle(sprite.x, sprite.y, sprite.width, sprite.height));
            }
            foreach (Sprite sprite in yellowFigures)
            {
                if (sprite.bmp != null && sprite.show == true)
                    g.DrawImage(sprite.bmp, new Rectangle(sprite.x, sprite.y, sprite.width, sprite.height));
            }
            foreach (Sprite sprite in blueFigures)
            {
                if (sprite.bmp != null && sprite.show == true)
                    g.DrawImage(sprite.bmp, new Rectangle(sprite.x, sprite.y, sprite.width, sprite.height));
            }
        }

        private void keyDown(object sender, KeyEventArgs e)
        {
            inkey = e.KeyCode.ToString();
        }

        private void keyUp(object sender, KeyEventArgs e)
        {
            inkey = "";
        }

        private void mouseClicked(object sender, MouseEventArgs e)
        {
            mouseKey = 1;
        }

        private void mouseDown(object sender, MouseEventArgs e)
        {
            mouseKey = 1;
        }

        private void mouseUp(object sender, MouseEventArgs e)
        {
            mouseKey = 0;
        }

        private void mouseMove(object sender, MouseEventArgs e)
        {
            mouseXp = e.X;
            mouseYp = e.Y;
        }

#endregion


        public static int ShowDialog(string text, string caption)
        {
            Form prompt = new Form()
            {
                Width = 200,
                Height = 150,
                FormBorderStyle = FormBorderStyle.FixedToolWindow,
                Text = caption,
                StartPosition = FormStartPosition.CenterScreen

            };
            Label textLabel = new Label() { Left = 20, Top = 23, Text = text };
            TextBox textBox = new TextBox() { Left = 105, Top = 20, Width = 55, TextAlign = HorizontalAlignment.Center };
            textBox.Text = "2";
            Button confirmation = new Button() { Text = "Potvrdi", Left = 20, Width = 140, Top = 70, DialogResult = DialogResult.OK };
            confirmation.Click += (sender, e) => { prompt.Close(); };
            prompt.Controls.Add(textBox);
            prompt.Controls.Add(confirmation);
            prompt.Controls.Add(textLabel);
            prompt.AcceptButton = confirmation;

            return prompt.ShowDialog() == DialogResult.OK ? Convert.ToInt16(textBox.Text) : 0;
        }
    }
}
