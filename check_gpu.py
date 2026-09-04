#!/usr/bin/env python3
"""
GPU Availability Checker for RPVE OCR System

This script checks if GPU acceleration is available for the OCR processing
and provides recommendations if it's not set up correctly.
"""

import sys

def check_gpu_availability():
    """Check GPU availability and provide detailed diagnostics."""
    
    print("=" * 70)
    print("RPVE OCR - GPU Availability Check")
    print("=" * 70)
    print()
    
    # 1. Check if PyTorch is installed
    try:
        import torch
        print("✅ PyTorch is installed")
        print(f"   Version: {torch.__version__}")
    except ImportError:
        print("❌ PyTorch is NOT installed")
        print("\n💡 Install PyTorch:")
        print("   pip install torch torchvision")
        return False
    
    print()
    
    # 2. Check CUDA availability
    cuda_available = torch.cuda.is_available()
    if cuda_available:
        print("✅ CUDA is available")
        print(f"   CUDA Version: {torch.version.cuda}")
    else:
        print("⚠️  CUDA is NOT available")
        print("\n💡 Possible reasons:")
        print("   1. No NVIDIA GPU in your system")
        print("   2. CUDA drivers not installed")
        print("   3. PyTorch installed without CUDA support (CPU-only version)")
        print("\n💡 To enable GPU:")
        print("   1. Check if you have NVIDIA GPU: Run 'nvidia-smi' in terminal")
        print("   2. Install CUDA-enabled PyTorch:")
        print("      pip uninstall torch torchvision")
        print("      pip install torch torchvision --index-url https://download.pytorch.org/whl/cu118")
        print("\n   Note: Your system will work fine on CPU, but GPU is 5-10x faster!")
        return False
    
    print()
    
    # 3. Check GPU details
    if cuda_available:
        try:
            gpu_count = torch.cuda.device_count()
            print(f"✅ Number of GPUs detected: {gpu_count}")
            print()
            
            for i in range(gpu_count):
                print(f"   GPU {i}:")
                print(f"      Name: {torch.cuda.get_device_name(i)}")
                
                props = torch.cuda.get_device_properties(i)
                total_memory = props.total_memory / 1024**3  # Convert to GB
                print(f"      Total Memory: {total_memory:.2f} GB")
                print(f"      Compute Capability: {props.major}.{props.minor}")
                
                # Check current memory usage
                if torch.cuda.is_initialized():
                    allocated = torch.cuda.memory_allocated(i) / 1024**3
                    reserved = torch.cuda.memory_reserved(i) / 1024**3
                    print(f"      Memory Allocated: {allocated:.2f} GB")
                    print(f"      Memory Reserved: {reserved:.2f} GB")
                    print(f"      Memory Free: {total_memory - allocated:.2f} GB")
                print()
                
        except Exception as e:
            print(f"⚠️  Error getting GPU details: {e}")
    
    # 4. Test GPU with a simple operation
    if cuda_available:
        try:
            print("🧪 Testing GPU with a simple operation...")
            device = torch.device("cuda")
            test_tensor = torch.zeros(1000, 1000).to(device)
            result = test_tensor @ test_tensor.T  # Matrix multiplication
            del test_tensor, result
            torch.cuda.empty_cache()
            print("✅ GPU test successful! GPU is working correctly.")
        except Exception as e:
            print(f"❌ GPU test failed: {e}")
            print("\n💡 Your GPU may not have enough memory or drivers have issues.")
            return False
    
    print()
    print("=" * 70)
    
    # 5. Check DocTR installation
    print("\n📦 Checking OCR dependencies...")
    try:
        from doctr.models import ocr_predictor
        print("✅ DocTR is installed")
    except ImportError:
        print("⚠️  DocTR is NOT installed")
        print("💡 Install: pip install python-doctr[torch]")
    
    print()
    
    # 6. Final summary
    if cuda_available:
        print("=" * 70)
        print("🎉 GPU ACCELERATION IS READY!")
        print("=" * 70)
        print("\nYour RPVE OCR system will automatically use GPU for:")
        print("  • 5-10x faster OCR processing")
        print("  • Batch processing of multiple pages")
        print("  • Reduced CPU load")
        print("\nThe system will automatically fall back to CPU if:")
        print("  • GPU runs out of memory")
        print("  • GPU encounters an error")
        print("  • You set RPVE_FORCE_CPU=true environment variable")
        return True
    else:
        print("=" * 70)
        print("💻 SYSTEM WILL USE CPU")
        print("=" * 70)
        print("\nYour system will work perfectly on CPU, but will be slower.")
        print("Expected processing time: ~40-50 seconds per 8-page document")
        print("\nWith GPU: ~5-10 seconds per 8-page document")
        return False

if __name__ == "__main__":
    success = check_gpu_availability()
    sys.exit(0 if success else 1)
