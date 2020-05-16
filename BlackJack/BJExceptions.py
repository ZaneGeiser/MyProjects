class InsufficientFundsError(Exception):
    """Raised When the Player has Insufficient funds for a bet"""
    pass

class CheatingBetError(Exception):
    """Raised when the player tries to bet 0 or nevative."""
    pass

class BustError(Exception):
    """Raised when the player or the dealer bust"""
    pass
