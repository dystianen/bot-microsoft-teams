const mongoose = require("mongoose");

const SuccessAccountSchema = new mongoose.Schema({
  email: { type: String, required: true },
  password: { type: String, required: true },
  telegram_id: { type: String, required: true },
  createdAt: { type: Date, default: Date.now },
});

const UserConfigSchema = new mongoose.Schema({
  telegram_id: { type: String, required: true, unique: true },
  microsoftUrl: { type: String, default: "https://admin.cloud.microsoft/?#/catalog/m/offer-details/microsoft-365-copilot/CFQ7TTC0MM8R" },
  concurrencyLimit: { type: Number, default: 5 },
  headless: { type: Boolean, default: false }, // Default window visible for user.
  updatedAt: { type: Date, default: Date.now },
});

module.exports = {
  SuccessAccount: mongoose.model("SuccessAccount", SuccessAccountSchema),
  UserConfig: mongoose.model("UserConfig", UserConfigSchema),
};
