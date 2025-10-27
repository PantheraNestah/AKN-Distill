import json
from channels.generic.websocket import AsyncWebsocketConsumer

class DocumentProcessingConsumer(AsyncWebsocketConsumer):
    async def connect(self):
        self.document_id = self.scope['url_route']['kwargs']['document_id']
        self.room_group_name = f'doc_{self.document_id}'

        # Join room group
        await self.channel_layer.group_add(
            self.room_group_name,
            self.channel_name
        )
        await self.accept()

    async def disconnect(self, close_code):
        # Leave room group
        await self.channel_layer.group_discard(
            self.room_group_name,
            self.channel_name
        )

    async def receive(self, text_data):
        # Handle any messages from client if needed
        pass

    async def processing_update(self, event):
        # Send processing update to WebSocket
        await self.send(text_data=json.dumps({
            'type': 'processing_update',
            'progress': event['progress'],
            'status': event['status'],
            'message': event['message']
        }))

    async def processing_complete(self, event):
        # Send completion message to WebSocket
        await self.send(text_data=json.dumps({
            'type': 'processing_complete',
            'preview_url': event['preview_url']
        }))