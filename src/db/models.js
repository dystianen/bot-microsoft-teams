const mongoose = require('mongoose');

const AccountHistorySchema = new mongoose.Schema({
  email: { type: String, required: true },
  password: { type: String, required: true },
  telegram_id: { type: String, required: true },
  status: { type: String, default: 'SUCCESS' }, // SUCCESS, FAILED
  log: { type: String },
  createdAt: { type: Date, default: Date.now },
});

const UserConfigSchema = new mongoose.Schema({
  telegram_id: { type: String, required: true, unique: true },
  microsoftUrl: {
    type: String,
    default:
      'https://admin.cloud.microsoft/?#/catalog/m/offer-details/microsoft-365-copilot/CFQ7TTC0MM8R',
  },
  concurrencyLimit: { type: Number, default: 5 },
  headless: { type: Boolean, default: false }, // Default window visible for user.
  updatedAt: { type: Date, default: Date.now },
});

module.exports = {
  AccountHistory: mongoose.model('AccountHistory', AccountHistorySchema),
  UserConfig: mongoose.model('UserConfig', UserConfigSchema),
};
