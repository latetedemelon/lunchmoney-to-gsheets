[![Donate](https://img.shields.io/badge/Donate-PayPal-green.svg)](https://paypal.me/latetedemelon) [![Donate](https://img.shields.io/badge/Donate-Buy%20Me%20a%20Coffee-yellow)](https://buymeacoffee.com/latetedemelon) [![Donate](https://img.shields.io/badge/Donate-Ko--Fi-ff69b4)](https://ko-fi.com/latetedemelon)

# Lunchmoney-to-Google Sheets

This script allows you to import data from the Lunchmoney.app API and present it in Google Sheets. You can retrieve data on categories, transactions, budgets, tags, Plaid accounts, assets, cryptocurrencies, and recurring expenses. The script will create a separate sheet for each endpoint.

## About

I had originally intended this to be a way to edit your data in Lunchmoney but doing this purely in Google Apps Script proved to be a real pain.  If there is enough interest I will pursue getting the update functionality working.

## Setup

1. Open a new Google Sheet.
2. Click on Extensions > Apps Script.
3. Cut and paste the provided `lunchmoney.gs` script, overwriting the existing content in the script editor.
4. Save the project and script names as desired.
5. Refresh your Google Sheet.
6. A new "Lunchmoney Configuration" sheet should be created. Add your API Key to the named field.
7. Update the date range for the data you're looking to retrieve.
8. A new Lunchmoney menu should now appear to the right of the Help menu.

## Usage

To retrieve data from Lunchmoney.app, select the desired option from the Lunchmoney menu. The script will create a separate sheet for each endpoint and update the data accordingly.

Available options:

- Refresh All: Updates data for all endpoints.
- Refresh Categories: Updates data for the `/categories` endpoint.
- Refresh Transactions: Updates data for the `/transactions` endpoint.
- Refresh Budgets: Updates data for the `/budgets` endpoint.
- Refresh Tags: Updates data for the `/tags` endpoint.
- Refresh Plaid Accounts: Updates data for the `/plaid_accounts` endpoint.
- Refresh Assets: Updates data for the `/assets` endpoint.
- Refresh Crypto: Updates data for the `/crypto` endpoint.
- Refresh Recurring Expenses: Updates data for the `/recurring_expenses` endpoint.

## Troubleshooting

If you encounter any issues, check the Apps Script editor's "View" > "Logs" for error messages or relevant information. Additionally, ensure that the API key and date ranges provided in the "Lunchmoney Configuration" sheet are correct.

## Contributing

Contributions to this project are welcome! If you have improvements, bug fixes, or new features you'd like to see added, please submit a Pull Request.

## Donations

If you find this project helpful and would like to support its development, consider making a donation to the project. Every little bit helps!

<a href='https://paypal.me/latetedemelon' target='_blank'><img src="https://github.com/stefan-niedermann/paypal-donate-button/blob/master/paypal-donate-button.png" width="270" height="105" alt='Donate via Paypay' />

<a href='https://ko-fi.com/latetedemelon' target='_blank'><img height='35' style='border:0px;height:46px;' src='https://az743702.vo.msecnd.net/cdn/kofi3.png?v=0' border='0' alt='Buy Me a Coffee at ko-fi.com' />

[!["Buy Me A Coffee"](https://www.buymeacoffee.com/assets/img/custom_images/yellow_img.png)](https://www.buymeacoffee.com/latetedemelon)

## License

This project is licensed under the MIT License. See the [LICENSE](https://github.com/latetedemelon/lunchmoney-to-gsheets/blob/main/LICENSE) file for details.
