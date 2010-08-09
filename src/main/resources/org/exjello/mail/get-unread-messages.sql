SELECT
    "urn:schemas:httpmail:fromemail"
FROM "" WHERE
    "DAV:iscollection" = False AND
    "DAV:ishidden" = False
AND "urn:schemas:httpmail:read" = False
ORDER BY "DAV:creationdate"
