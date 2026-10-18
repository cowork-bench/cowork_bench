"""Setup failures that must not be charged to an agent's task score."""


class PreprocessError(RuntimeError):
    """Task preparation failed before the agent could execute the task."""
