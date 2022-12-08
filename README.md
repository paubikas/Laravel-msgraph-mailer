# Laravel Microsoft Graph Mail

This package makes it easy to send emails from your personal, work or school account using Microsoft's Graph API,
allowing you to benefit from HTTP instead of SMTP with Laravel.

Inspired by wapacro/laravel-msgraph-mail and fixed to work with Laravel 9
## Installation

Install the package using composer:

```
composer require paubikas/laravel-msgraph-mailer
```

Add the configuration to your mail.php config file:

```php
'mailers' => [
    'microsoft-graph' => [
        'transport'       => 'microsoft-graph',
        'tenant'          => env('MAIL_MSGRAPH_TENANT', 'common'),
        'client'          => env('MAIL_MSGRAPH_CLIENT'),
        'secret'          => env('MAIL_MSGRAPH_SECRET')
        'saveToSentItems' => env('MAIL_MSGRAPH_SAVE_TO_SENT_ITEMS', false)
    ]
    // ...
]
```

Add the configuration to your ENV file:

```env
MAIL_MAILER=microsoft-graph
MAIL_MSGRAPH_TENANT=
MAIL_MSGRAPH_CLIENT=
MAIL_MSGRAPH_SECRET=
MAIL_MSGRAPH_SAVE_TO_SENT_ITEMS=false
```

Valid values for `tenant` are your tenant identifier (work & school accounts) or `common` for personal accounts.

**Note:** This package relies on [Laravel's Cache](https://laravel.com/docs/cache) interface for caching access tokens.
Make sure to configure it properly, too!


### Getting the credentials

To get the necessary client ID and secret you'll need to register your application and grant it the required
permissions. Head over to [the Azure Portal to do so](https://docs.microsoft.com/en-us/graph/auth-register-app-v2)
(you don't need to be an Azure user).

Make sure to grant the _Mail.Send_ permission and to generate a secret afterwards (may be hidden during app registration).

**Work & School accounts:** Granting your app the _Mail.Send_ permission allows you by default to send emails with every
valid email address within your company/school. Use an [Exchange Online Application Access Policy](https://docs.microsoft.com/en-us/graph/auth-limit-mailbox-access)
to restrict which email addresses are valid senders for your application.
