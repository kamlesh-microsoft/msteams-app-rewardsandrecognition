## Assumptions

The estimate below assumes:

-   500 users in the tenant
-   Each user performs 5 add, update or delete operations per day.
-   Each user uses messaging extension 25 times/week.

## [](/wiki/costestimate#sku-recommendations)SKU recommendations

The recommended SKUs for a production environment are:

-   App Service: Standard (S1)
-   Azure Search: Basic
    -   The Azure Search service cannot be upgraded once it is provisioned, so select a tier that will meet your anticipated needs.

## Estimated load

**Data storage**: 1 GB max    

**Table data operations**:

* Bot
	* Total number of read calls in storage = 4 calls/hour * 24 hours/day * 30 days = 2880
	* Total number of write calls in storage =  4 calls/hour * 24 hours/day * 30 days = 2880 

## Estimated cost

**IMPORTANT:** This is only an estimate, based on the assumptions above. Your actual costs may vary.

Prices were taken from the [Azure Pricing Overview](https://azure.microsoft.com/en-us/pricing/) on 08 April 2020, for the West US 2 region.

Use the [Azure Pricing Calculator](https://azure.com/e/70ac9cd54e3841999610b06e019d9b68) to model different service tiers and usage patterns.

Resource                                    | Tier          | Load          | Monthly price
---                                         | ---           | ---           | --- 
Storage account (Table)                     | Standard_LRS  | < 1GB data, 5,760 operations | $0.05 + $0.01 = $0.06
Bot Channels Registration                   | F0            | N/A           | Free
App Service Plan                            | S1            | 744 hours     | $74.40
App Service (Messaging Extension)           | -             |               | (charged to App Service Plan) 
Application Insights (Messaging Extension)  | -             | < 5GB data    | (free up to 5 GB)
Azure Search                                | B             |               | $75.14
**Total**                                   |               |               | **$149.6**
