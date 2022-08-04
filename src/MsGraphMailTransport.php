<?php

namespace LaravelMsGraphMailer;

use GuzzleHttp\Exception\BadResponseException;
use GuzzleHttp\Exception\ConnectException;
use Illuminate\Support\Facades\Cache;
use LaravelMsGraphMailer\Exceptions\CouldNotGetToken;
use LaravelMsGraphMailer\Exceptions\CouldNotReachService;
use LaravelMsGraphMailer\Exceptions\CouldNotSendMail;
use Psr\EventDispatcher\EventDispatcherInterface;
use Psr\Log\LoggerInterface;
use Symfony\Component\HttpClient\HttpClient;
use Symfony\Component\Mailer\Envelope;
use Symfony\Component\Mailer\SentMessage;
use Symfony\Component\Mailer\Transport\AbstractApiTransport;
use Symfony\Component\Mime\Address;
use Symfony\Component\Mime\Email;
use Symfony\Contracts\HttpClient\HttpClientInterface;
use Symfony\Contracts\HttpClient\ResponseInterface;

class MsGraphMailTransport extends AbstractApiTransport
{
    /**
     * @var string
     */
    protected string $tokenEndpoint = 'https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token';

    /**
     * @var string
     */
    protected string $apiEndpoint = 'https://graph.microsoft.com/v1.0/users/{from}/sendMail';


    private string $secret;
    private string $tenant_id;
    private ?string $client_id;

    protected $http;

    public function __construct($config, HttpClientInterface $client = null, EventDispatcherInterface $dispatcher = null, LoggerInterface $logger = null)
    {
        $this->secret = $config['secret'];
        $this->tenant_id = $config['tenant'];
        $this->client_id = $config['client'];
        $this->saveToSentItems = $config['saveToSentItems'];
        $this->http = $client ?? HttpClient::create();

        parent::__construct($client, $dispatcher, $logger);
    }

    protected function doSendApi(SentMessage $sentMessage, Email $email, Envelope $envelope): ResponseInterface {
//        $this->beforeSendPerformed($message);
        $rawPayload = $this->getPayload($email, $envelope);

        $url = str_replace('{from}', urlencode($envelope->getSender()->getAddress()), $this->apiEndpoint);

        $response = $this->http->request('POST', $url, [
            'headers' => $this->getHeaders(),
            'json' => [
                'message' => $rawPayload,
                'saveToSentItems' => ($this->saveToSentItems) ? $this->saveToSentItems : true
            ]
        ]);

        try {
            $statusCode = $response->getStatusCode();

        } catch (BadResponseException $e) {
            // The API responded with 4XX or 5XX error
            if ($e->hasResponse()) $response = json_decode((string)$e->getResponse()->getBody());
            throw CouldNotSendMail::serviceRespondedWithError($response->error->code ?? 'Unknown', $response->error->message ?? 'Unknown error');
        } catch (ConnectException $e) {
            // A connection error (DNS, timeout, ...) occurred
            throw CouldNotReachService::networkError();
        } catch (Throwable $e) {
            throw CouldNotReachService::unknownError();
        }
        return $response;
    }

    public function __toString(): string {
        return $this->apiEndpoint;
    }


    /**
     * Transforms given SwiftMailer message instance into
     * Microsoft Graph message object
     * @param Email $message
     * @param Envelope $envelope
     * @return array
     */
    protected function getPayload(Email $email, Envelope $envelope): array {
        $from = $envelope->getSender();
        $priority = $email->getPriority();
        $html = $email->getHtmlBody();

        [$attachments, $html] = $this->prepareAttachments($email, $html);

        return array_filter([
            'subject' => $email->getSubject(),
            'sender' => $this->toRecipientCollection([$from])[0],
            'from' => $this->toRecipientCollection([$from])[0],
            'replyTo' => $this->toRecipientCollection($email->getReplyTo()),
            'toRecipients' => $this->toRecipientCollection($this->getRecipients($email, $envelope)),
            'ccRecipients' => $this->toRecipientCollection($email->getCc()),
            'bccRecipients' => $this->toRecipientCollection($email->getBcc()),
            'importance' => $priority === 3 ? 'Normal' : ($priority < 3 ? 'Low' : 'High'),
            'body' => [
                'contentType' => 'html',
                'content' => $html,
            ],
            'attachments' => $attachments,
        ]);
    }

    /**
     * Transforms given SimpleMessage recipients into
     * Microsoft Graph recipients collection
     * @param array|string $recipients
     * @return array
     */
    protected function toRecipientCollection($recipients): array {
        $collection = [];

        // If the provided list is empty
        // return an empty collection
        if (!$recipients) {
            return $collection;
        }

        // Some fields yield single e-mail
        // addresses instead of arrays
        if (is_string($recipients)) {
            $collection[] = [
                'emailAddress' => [
                    'name' => null,
                    'address' => $recipients,
                ],
            ];

            return $collection;
        }

        foreach($recipients as $recipientKey => $recipient) {
            if($recipient instanceof Address) {
                $collection[] = [
                    'emailAddress' => [
                        'name' => $recipient->getName(),
                        'address' => $recipient->getAddress()
                    ]
                ];
            } else {
                $collection[] = [
                    'emailAddress' => [
                        'name' => $recipient,
                        'address' => $recipientKey,
                    ],
                ];
            }
        }
        return $collection;
    }

    private function prepareAttachments(Email $email, ?string $html): array
    {
        $attachments = $inlines = [];
        foreach ($email->getAttachments() as $attachment) {
            $headers = $attachment->getPreparedHeaders();
            if ('inline' === $headers->getHeaderBody('Content-Disposition')) {
                // replace the cid with just a file name (the only supported way by Mailgun)
                if ($html) {
                    $filename = $headers->getHeaderParameter('Content-Disposition', 'filename');
                    $new = basename($filename);
                    $html = str_replace('cid:'.$filename, 'cid:'.$new, $html);
                    $p = new \ReflectionProperty($attachment, 'filename');
                    $p->setAccessible(true);
                    $p->setValue($attachment, $new);
                    $attachments[] = [
                        "@odata.type" => "#microsoft.graph.fileAttachment",
                        "name" => $headers->getHeaderParameter('Content-Type', 'name'),
                        "contentType" => $headers->getHeaderBody('Content-Type'),
                        "contentBytes" => $attachment->bodyToString(),
                        "contentId" => $new,
                        "isInline" => true
                    ];
                }
                $inlines[] = $attachment;

            } else {
                $attachments[] = [
                    "@odata.type" => "#microsoft.graph.fileAttachment",
                    "name" => $headers->getHeaderParameter('Content-Type', 'name'),
                    "contentType" => $headers->getHeaderBody('Content-Type'),
                    "contentBytes" => $attachment->bodyToString()
                ];
            }
        }

        return [$attachments, $html];
    }

    /**
     * Transforms given SwiftMailer children into
     * Microsoft Graph attachment collection
     * @param $attachments
     * @return array
     */
    protected function toAttachmentCollection($attachments): array {
        $collection = [];

        foreach ($attachments as $attachment) {
            if (!$attachment instanceof Swift_Mime_Attachment) {
                continue;
            }

            $collection[] = [
                'name' => $attachment->getFilename(),
                'contentId' => $attachment->getId(),
                'contentType' => $attachment->getContentType(),
                'contentBytes' => base64_encode($attachment->getBody()),
                'size' => strlen($attachment->getBody()),
                '@odata.type' => '#microsoft.graph.fileAttachment',
                'isInline' => $attachment instanceof Swift_Mime_EmbeddedFile,
            ];

        }

        return $collection;
    }

    /**
     * Returns header collection for API request
     * @return string[]
     * @throws CouldNotGetToken
     * @throws CouldNotReachService
     */
    protected function getHeaders(): array {
        return [
            'Accept' => 'application/json',
            'Content-Type' => 'application/json',
            'Authorization' => 'Bearer ' . $this->getAccessToken(),
        ];
    }

    /**
     * Returns API access token
     * @return string
     * @throws CouldNotReachService
     * @throws CouldNotGetToken
     */
    protected function getAccessToken(): string {
        try {
            return Cache::remember('mail-msgraph-accesstoken', 45, function () {
                $url = str_replace('{tenant}', $this->tenant_id ?? 'common', $this->tokenEndpoint);
                $response = $this->http->request('POST', $url, [
                    'body' => [
                        'client_id' => $this->client_id,
                        'client_secret' => $this->secret,
                        'scope' => 'https://graph.microsoft.com/.default',
                        'grant_type' => 'client_credentials',
                    ],
                ]);
                $response = $response->toArray();
                return $response['access_token'];
            });
        } catch (BadResponseException $e) {
            // The endpoint responded with 4XX or 5XX error
            $response = json_decode((string)$e->getResponse()->getBody());
            throw CouldNotGetToken::serviceRespondedWithError($response->error, $response->error_description);
        } catch (ConnectException $e) {
            // A connection error (DNS, timeout, ...) occurred
            throw CouldNotReachService::networkError();
        } catch (Throwable $e) {
            // An unknown error occurred
            throw CouldNotReachService::unknownError();
        }
    }

}
