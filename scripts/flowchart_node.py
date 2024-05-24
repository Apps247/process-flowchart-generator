class FlowchartNode:
    def __init__(self, id, visioId, type, title, description, actor):
        self.id = id
        self.visioId = visioId
        self.type = type
        self.title = title
        self.description = description
        self.actor = actor