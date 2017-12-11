# InPlace E-Discovery Search and Purge - Exchange Online
## Ediscovery search and purge functionality that actually work!
*Search **all** mailboxes, get **all** results despite undisclosed throttle and caps*

Ediscovery on Exchange Online is throttled and capped.  Large organizations require precise and complete search results against large sets of mailboxes.  The tools provided by Microsoft return only 1000 results against a maximum of 20,000 mailboxes.  There are additional caps on these searches but combine that with throttling and timeouts, it becomes almost unusable.  There are many combinations of scenarios which Microsoft ediscovery tools just will not suffice.  If you've used the solution on the admin portal, you're in for a really bad experience of a long running job requiring use of Internet Explorer, which cannot be automated.

These functions here are resilient to throttling and timeouts.  It works within the caps imposed and provides all search results despite the number of mailboxes your organization has.  Once results are returned, the emails can be examined and/or purged.

The functions built here were originally created to enforce custom email retention rules not provided by Exchange or Exchange Online.  We also found utilities here useful recently to eradicate mail from a targeted phishing attack.

**Prerequisites:**
1. provide your own KQL (keyword query languange) string.  sender/recipient fields are indexed, so EXO ignores the @domain component..
2. have ApplicationImpersonation rights against all your mailboxes
3. have a dynamic distribution list that contains all mailboxes you wish to target (mine is hidden too)


To get your complete search result:
```
$Results = Start-FullCustomEdiscoverySearch -service $EWSservice -KQL $YourKqlString 
   -MailboxGroupFilter $MyTargetMailboxes
```
  *oh, btw... EXO only allows up to 2 concuurent ediscovery searches.  be advised.*
