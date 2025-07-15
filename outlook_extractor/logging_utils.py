"""Logging utilities for Outlook Extractor.

This module provides additional logging utilities and decorators
for the application.
"""
import functools
import inspect
import logging
import time
from contextlib import contextmanager
from typing import Any, Callable, Dict, Optional, TypeVar, cast

from outlook_extractor.logging_config import get_logger, LogErrors

# Type variable for generic function type
F = TypeVar('F', bound=Callable[..., Any])

def log_function_call(log_args: bool = True, log_result: bool = False, level: int = logging.DEBUG):
    """
    Decorator to log function calls and their results.
    
    Args:
        log_args: Whether to log function arguments
        log_result: Whether to log the function result
        level: Logging level to use
    """
    def decorator(func: F) -> F:
        logger = get_logger(func.__module__)
        
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            # Log function entry
            if log_args:
                bound_args = inspect.signature(func).bind(*args, **kwargs)
                bound_args.apply_defaults()
                logger.log(
                    level,
                    "Calling %s(%s)",
                    func.__qualname__,
                    ", ".join(f"{k}={v!r}" for k, v in bound_args.arguments.items())
                )
            else:
                logger.log(level, "Calling %s", func.__qualname__)
            
            # Call the function
            start_time = time.monotonic()
            try:
                result = func(*args, **kwargs)
                
                # Log the result if requested
                if log_result:
                    logger.log(level, "%s -> %r", func.__qualname__, result)
                
                return result
                
            except Exception as e:
                # Log the exception
                logger.exception(
                    "%s raised %s: %s",
                    func.__qualname__,
                    type(e).__name__,
                    str(e)
                )
                raise
            finally:
                # Log the execution time
                duration = time.monotonic() - start_time
                logger.log(
                    level,
                    "%s completed in %.3f seconds",
                    func.__qualname__,
                    duration
                )
        
        return cast(F, wrapper)
    return decorator

@contextmanager
def log_duration(description: str, level: int = logging.INFO, logger: Optional[logging.Logger] = None):
    """
    Context manager to log the duration of a code block.
    
    Args:
        description: Description of the operation being timed
        level: Logging level to use
        logger: Logger to use (defaults to module logger)
    """
    if logger is None:
        # Get the logger of the calling function's module
        frame = inspect.currentframe()
        try:
            if frame is not None and frame.f_back is not None and frame.f_back.f_globals is not None:
                module_name = frame.f_back.f_globals.get('__name__', __name__)
            else:
                module_name = __name__
        finally:
            # Avoid reference cycles
            del frame
    
    logger = logger or get_logger(module_name)
    start_time = time.monotonic()
    
    try:
        logger.log(level, "%s - started", description)
        yield
    except Exception as e:
        logger.exception(
            "%s - failed after %.3f seconds: %s",
            description,
            time.monotonic() - start_time,
            str(e)
        )
        raise
    else:
        logger.log(
            level,
            "%s - completed in %.3f seconds",
            description,
            time.monotonic() - start_time
        )

class LogContext:
    """Context manager for logging with additional context information."""
    
    def __init__(self, logger: logging.Logger, msg: str, *args: Any, **kwargs: Any):
        self.logger = logger
        self.msg = msg
        self.args = args
        self.kwargs = kwargs
        self.start_time: Optional[float] = None
    
    def __enter__(self):
        self.start_time = time.monotonic()
        self.logger.info("%s - started", self.msg, *self.args, **self.kwargs)
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        duration = time.monotonic() - (self.start_time or 0)
        if exc_val is not None:
            self.logger.error(
                "%s - failed after %.3f seconds: %s",
                self.msg,
                duration,
                str(exc_val),
                *self.args,
                exc_info=(exc_type, exc_val, exc_tb),
                **self.kwargs
            )
        else:
            self.logger.info(
                "%s - completed in %.3f seconds",
                self.msg,
                duration,
                *self.args,
                **self.kwargs
            )
        return False  # Don't suppress the exception

def log_errors(logger: Optional[logging.Logger] = None, **kwargs):
    """
    Decorator to log exceptions raised by a function.
    
    Args:
        logger: Logger to use (defaults to module logger)
        **kwargs: Additional keyword arguments to pass to the logger
    """
    def decorator(func: F) -> F:
        nonlocal logger
        if logger is None:
            logger = get_logger(func.__module__)
        
        @functools.wraps(func)
        def wrapper(*args, **kw):
            with LogErrors(logger, f"Error in {func.__qualname__}", **kwargs):
                return func(*args, **kw)
        
        return cast(F, wrapper)
    return decorator

def log_async_errors(logger: Optional[logging.Logger] = None, **kwargs):
    """
    Decorator to log exceptions raised by an async function.
    
    Args:
        logger: Logger to use (defaults to module logger)
        **kwargs: Additional keyword arguments to pass to the logger
    """
    def decorator(func: F) -> F:
        nonlocal logger
        if logger is None:
            logger = get_logger(func.__module__)
        
        @functools.wraps(func)
        async def wrapper(*args, **kw):
            try:
                return await func(*args, **kw)
            except Exception as e:
                logger.error(
                    f"Error in {func.__qualname__}: {e}",
                    exc_info=True,
                    **kwargs
                )
                raise
        
        return cast(F, wrapper)
    return decorator
