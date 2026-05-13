# G3AXONLIVE Properties

The following properties are available for the `G3AXONLIVE` object.

| Property | Access | Type | Description |
|---|---|---|---|
| EventArgs | Read-only | String | Returns the entire event arguments map encoded as a JSON string. |
| EventComponentID | Read-only | String | Returns the ID of the component that fired the asynchronous event. |
| EventName| Read-only | String | Returns the name of the asynchronous event that was fired (e.g., `onclick`). |
| IsAsyncRequest | Read-only | Boolean | Returns `True` if the current request is an asynchronous G3AxonLive event POST, otherwise `False`. |
| Version | Read-only | String | Returns the current version of the AxonLive framework (e.g., `2.0.0`). |
