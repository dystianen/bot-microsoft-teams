const mongoose = require("mongoose");

const SuccessAccountSchema = new mongoose.Schema({
  email: { type: String, required: true },
  password: { type: String, required: true },
  domainEmail: { type: String },
  domainPassword: { type: String },
  telegram_id: { type: String, required: true },
  createdAt: { type: Date, default: Date.now },
});

const VCCSchema = new mongoose.Schema({
  cardNumber: { type: String, required: true, unique: true },
  cvv: { type: String, required: true },
  expMonth: { type: String, required: true },
  expYear: { type: String, required: true },
  saldo: { type: Number, default: 3 },
  status: { type: String, default: "active" }, // active, empty
  telegram_id: { type: String, required: true },
  createdAt: { type: Date, default: Date.now },
});

const UserConfigSchema = new mongoose.Schema({
  telegram_id: { type: String, required: true, unique: true },
  microsoftUrl: { type: String, default: "https://signup.microsoft.com/get-started/signup?products=2ef55228-927f-4abc-ac51-befd4cdcd850&mproducts=CFQ7TTC0LF8S:0009&fmproducts=CFQ7TTC0LF8S:0009&renewalterm=P1Y&renewalbillingterm=P1Y&culture=ph-ph&country=ph&ali=1" },
  concurrencyLimit: { type: Number, default: 2 },
  maxAccountsPerPayment: { type: Number, default: 3 },
  proxyUsername: { type: String, default: "" },
  proxyPassword: { type: String, default: "" },
  stopPoint: { type: String, default: "full" }, // "vcc_success" or "full"
  updatedAt: { type: Date, default: Date.now },
});

module.exports = {
  SuccessAccount: mongoose.model("SuccessAccount", SuccessAccountSchema),
  VCC: mongoose.model("VCC", VCCSchema),
  UserConfig: mongoose.model("UserConfig", UserConfigSchema),
};
