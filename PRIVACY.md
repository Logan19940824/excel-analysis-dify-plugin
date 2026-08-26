# Privacy

The plugin receives the Dify file URL and downloads the referenced Excel or CSV file for analysis. It processes the file in a temporary local database, returns requested metadata or query results, and removes temporary files after each invocation.

The plugin does not intentionally retain workbooks, query history, API keys, or user profile data. The file URL and file contents are sent only to the URL supplied by Dify and are not sent to another third party by this plugin.
