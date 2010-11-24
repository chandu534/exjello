SELECT
    "urn:schemas:httpmail:fromemail"
FROM "" WHERE
    "DAV:iscollection" = False AND
    "DAV:ishidden" = False
{BOOKMARK_FILTER_UNREADED}
{BOOKMARK_FILTER_TO}
{BOOKMARK_FILTER_FROM}
{BOOKMARK_FILTER_NOT_FROM}
{BOOKMARK_FILTER_LAST_CHECK}
ORDER BY "DAV:creationdate"