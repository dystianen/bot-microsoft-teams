const TeamsBot = require('./bots/TeamsBot');

async function processSingleAccount(accountConfig, index, total) {
  console.log(
    `\n--- Starting Account ${index + 1} of ${total}: ${accountConfig.microsoftAccount.email} ---`
  );

  let bot = null;
  let executionResult = null;

  try {
    // If you want to use AdsPower, you'd need to add that logic back here.
    // For now, we launch a local browser (incognito) as requested.
    bot = new TeamsBot(null, accountConfig);
    const result = await bot.run();

    if (result && result.success) {
      console.log(`[Account ${index + 1}] Automation finished successfully.`);
      executionResult = {
        status: 'SUCCESS',
        log: 'Completed successfully',
      };
    } else {
      console.error(
        `[Account ${index + 1}] Automation failed: ${result?.error || 'Unknown error'}`
      );
      executionResult = {
        status: 'FAILED',
        log: result?.error || 'Unknown automation error',
      };
    }
  } catch (err) {
    console.error(`\n[ERROR Account ${index + 1}] failed:`, err.message);
    executionResult = {
      status: 'FAILED',
      log: err.message,
    };
  } finally {
    if (bot) {
      try {
        await bot.cleanup();
      } catch (e) {
        console.error(`[Account ${index + 1}] Bot cleanup error:`, e.message);
      }
    }
  }

  return executionResult;
}

module.exports = {
  processSingleAccount,
};
