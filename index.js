const adsPowerHelper = require("./adspower_helper");
const TeamsBot = require("./teams_bot");

async function main() {
  const accountConfig = {
    microsoftAccount: {
      email: "miacampbell@capitalhrgroup.onmicrosoft.com",
      password: "MiaHR424"
    },
    // Add other config here as needed
  };

  let profileId = null;
  try {
    console.log("Creating profile...");
    profileId = await adsPowerHelper.createProfile("TeamsBotProfile");
    
    console.log("Starting browser...");
    const { wsUrl } = await adsPowerHelper.startBrowser(profileId);
    
    console.log("Running bot...");
    const bot = new TeamsBot(wsUrl, accountConfig);
    await bot.run();
    
  } catch (error) {
    console.error("Main execution failed:", error.message);
  } finally {
    // Cleanup if needed
  }
}

if (require.main === module) {
  main();
}
