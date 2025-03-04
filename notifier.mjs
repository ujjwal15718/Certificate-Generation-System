import gTTS from "gtts";
import chalk from "chalk";
import figlet from "figlet";
import { exec } from "child_process";

// Function to speak text using gTTS
function speak(text) {
    const gtts = new gTTS(text, "en");
    gtts.save("voice.mp3", (err) => {
        if (err) {
            console.error("❌ Text-to-Speech Error:", err);
        } else {
            exec("mpg321 voice.mp3", (error) => {
                if (error) console.error("❌ Error playing audio:", error);
            });
        }
    });
}

// Function to display a stylish banner
export function showBanner() {
    console.log(chalk.blue.bold(figlet.textSync("NavGurukul", { horizontalLayout: "full" })));
    console.log(chalk.green.bold("Meraki Certificate Generator"));
    console.log(chalk.yellow("By Developer Ujjwal"));
    console.log(chalk.magenta("=========================================="));
}

// Function to announce certificate generation (ONLY PRINTS, NO SPEAKING)
export function announceStart() {
    console.log(chalk.cyan("🎙️ Your certificate is generating, please be patient..."));
}

// Function to announce successful generation (SPEAKS OUT LOUD)
export function announceSuccess() {
    console.log(chalk.green("✅ Generated successfully! 🎉"));
    speak("Generated successfully!");
}
