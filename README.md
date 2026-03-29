# gift-card-balance-checker
Google Apps Script for automated gift card balance checking with CAPTCHA handling. Features: window reuse, queueing system, progress notifications, and seamless batch processing.

## Supported Gift Card Types

This script supports two types of gift cards:

### 1. Standard Gift Cards
- **Format:** Card Number, CVV, Expiry Date
- **Balance Check URL:** https://cardbalance.com.au/ (default)
- **Examples:** Coles, Woolworths, Bunnings, and most other retailers

### 2. David Jones Reward Cards
- **Format:** Reward Number, PIN, Expiry Date
- **Balance Check URL:** https://www.davidjones.com/rewards/balance-check
- **Auto-detection:** The script automatically detects David Jones cards when "David Jones" is in the Retailer column (Column B)
- **Display:** Shows "Reward Number" and "PIN" instead of "Card Number" and "CVV"

## Spreadsheet Column Structure

| Column | Field | Description |
|--------|-------|-------------|
| A | Date Added | When the card was added to tracker |
| B | Retailer | Store name (e.g., "David Jones", "Coles") |
| C | Card Type | Type of gift card |
| D | Card Number | Card/Reward number |
| E | Card Value | Original card value |
| F | CVV/PIN | Security code or PIN |
| G | Additional Info | Any extra notes |
| H | Expiry Date | Card expiration date |
| I | Balance URL | Custom balance check URL (optional) |
| J+ | Current Balance | Auto-populated by script |

## How It Works

1. **Add Your Cards:** Enter card details in the spreadsheet
2. **For David Jones:** Enter "David Jones" in the Retailer column (Column B), the Reward Number in Column D, and PIN in Column F
3. **Run Check:** Use the menu: 🧧 Card Balance → Check All Balances
4. **Handle CAPTCHA:** Complete any CAPTCHA challenges in the popup window
5. **Enter Balance:** After verification, manually enter the balance shown
6. **Auto-Continue:** Script automatically moves to the next card

## Features

- ✅ **Multi-retailer Support:** Standard gift cards + David Jones Rewards
- ✅ **Smart Detection:** Automatically detects card type based on retailer name
- ✅ **Batch Processing:** Check multiple cards with queueing system
- ✅ **CAPTCHA Handling:** Opens balance check pages for manual CAPTCHA completion
- ✅ **Progress Tracking:** Toast notifications show current progress
- ✅ **Window Reuse:** Efficient processing with single popup window
- ✅ **Resume Capability:** Continue after CAPTCHA interruptions
