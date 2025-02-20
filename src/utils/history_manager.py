class HistoryManager:
    """历史记录管理器"""
    def __init__(self):
        self.history = []
        self.redo_stack = []
        self.initial_state = None
        
    def save_state(self, state):
        """保存状态"""
        self.history.append(state)
        self.redo_stack.clear()
        
        if self.initial_state is None:
            self.initial_state = state.copy()
            
    def can_undo(self):
        return len(self.history) > 1
        
    def can_redo(self):
        return len(self.redo_stack) > 0
        
    def can_reset(self):
        return self.initial_state is not None 

    def undo(self):
        """撤销操作"""
        if self.can_undo():
            state = self.history.pop()
            self.redo_stack.append(state)
            return self.history[-1]
        return None

    def redo(self):
        """重做操作"""
        if self.can_redo():
            state = self.redo_stack.pop()
            self.history.append(state)
            return state
        return None

    def reset(self):
        """重置状态"""
        if self.can_reset():
            state = self.initial_state.copy()
            self.history.clear()
            self.redo_stack.clear()
            self.history.append(state)
            return state
        return None 