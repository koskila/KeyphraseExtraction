# Keyword and Tagging examples

This is my public repository for testing KeyPhrase extraction and application.

## KeywordWebhookReceiver

An Azure Function that takes input from a Flow (for example), fetches an item from SharePoint list, extracts the text, analyzes it for keyphrases (using Azure Cognitive Services) and then writes the keyphrases to the list item's taxonomy field.

### Getting Started

Coming later.

### Demo

For a demo of the KeywordWebhookReceiver, see this video: https://youtu.be/G0kESOlBBjk
For a blog post about the core idea, see this: https://www.koskila.net/2018/01/12/resolve-managed-metadata-madness-sharepoint/ 

#### Prerequisites

- Azure subscription (you'll create the Azure Function there)
- SharePoint Online
- Azure Cognitive Services API key
- Flow (even the free plan should do, but with SharePoint you should have an Office 365 license, so even better!)

## Koskila.KeywordManager

Retired example using Azure Cognitive Services Topics -API

### Demo

Not available, retired example code.