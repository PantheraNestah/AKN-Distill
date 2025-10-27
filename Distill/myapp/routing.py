from django.urls import re_path
from . import consumers

websocket_urlpatterns = [
    re_path(r'ws/process/(?P<document_id>\d+)/$', consumers.DocumentProcessingConsumer.as_asgi()),
]