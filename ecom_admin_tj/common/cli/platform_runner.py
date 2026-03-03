"""
Platform Runner - Generic CLI entry point for platform processing

Provides standardized way to run admin processing for any platform
(Shopee, Lazada, Tiktok) with consistent error handling.
"""
import sys
from typing import Type, Callable, Optional
from ..base import Base


class PlatformRunner:
    """Generic runner for platform admin processing"""
    
    @staticmethod
    def run(
        platform_class: Type[Base],
        suppress_warnings: bool = False,
        error_handler: Optional[Callable[[Exception], None]] = None,
        exit_on_error: bool = True
    ) -> Optional[Base]:
        """
        Generic entry point for platform processing
        
        Args:
            platform_class: Platform class (Shopee, Lazada, Tiktok)
            suppress_warnings: Whether to suppress openpyxl/pandas warnings
            error_handler: Custom error handler function
            exit_on_error: Whether to exit on error (default: True for CLI)
            
        Returns:
            Platform instance if successful, None if error occurred
            
        Example:
            ```python
            # In shopee/__main__.py
            from ..common.cli.platform_runner import PlatformRunner
            from .shopee import Shopee
            
            if __name__ == "__main__":
                PlatformRunner.run(Shopee)
            ```
        """
        if suppress_warnings:
            import warnings
            warnings.filterwarnings("ignore", category=UserWarning)
            warnings.filterwarnings("ignore", category=FutureWarning)
        
        try:
            # Create instance from command-line arguments
            instance = platform_class.from_args()
            
            # Run processing
            instance.process()
            
            return instance
            
        except ValueError as e:
            # Handle expected errors (invalid arguments, data issues)
            if error_handler:
                error_handler(e)
            else:
                print(f"❌ Error: {e}")
            
            if exit_on_error:
                sys.exit(1)
            return None
            
        except FileNotFoundError as e:
            # Handle file not found errors
            if error_handler:
                error_handler(e)
            else:
                print(f"❌ File not found: {e.filename if hasattr(e, 'filename') else e}")
            
            if exit_on_error:
                sys.exit(1)
            return None
            
        except KeyboardInterrupt:
            print("\n⚠️  Process interrupted by user")
            if exit_on_error:
                sys.exit(130)  # Standard exit code for SIGINT
            return None
            
        except Exception as e:
            # Handle unexpected errors
            if error_handler:
                error_handler(e)
            else:
                print(f"❌ Unexpected error: {e}")
                import traceback
                traceback.print_exc()
            
            if exit_on_error:
                sys.exit(1)
            return None
