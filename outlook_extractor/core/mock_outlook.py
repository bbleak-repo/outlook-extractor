class MockOutlookClient:
    def __init__(self, *args, **kwargs):
        raise RuntimeError("Outlook integration is only available on Windows")
