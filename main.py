import argparse
import os
from powerpoint_analyzer import analyze_powerpoint, GeminiModel

def main():
    parser = argparse.ArgumentParser(description="Analyze PowerPoint presentations using AI models.")
    parser.add_argument("input", help="Path to the input PowerPoint file")
    parser.add_argument("output", help="Path to the output text file")
    parser.add_argument("--model", choices=["gemini"], default="gemini", help="AI model to use (default: gemini)")
    parser.add_argument("--combine-images", action="store_true", help="Combine images from slides")
    parser.add_argument("--no-labels", action="store_false", dest="add_labels", help="Don't add labels to images")
    parser.add_argument("--api-key", help="API key for the AI model")

    args = parser.parse_args()

    # Validate input file
    if not os.path.isfile(args.input):
        print(f"Error: Input file '{args.input}' does not exist.")
        return

    # Get API key
    api_key = args.api_key or os.environ.get("GEMINI_API_KEY")
    if not api_key:
        print("Error: API key not provided. Use --api-key or set the GEMINI_API_KEY environment variable.")
        return

    # Initialize the AI model
    if args.model == "gemini":
        ai_model = GeminiModel(api_key=api_key)
    else:
        print(f"Error: Unsupported model '{args.model}'")
        return

    # Run the analysis
    try:
        analyze_powerpoint(
            pptx_file=args.input,
            output_file=args.output,
            ai_model=ai_model,
            combine_images=args.combine_images,
            add_labels=args.add_labels
        )
        print(f"Analysis complete. Results saved to {args.output}")
    except Exception as e:
        print(f"An error occurred during analysis: {str(e)}")

if __name__ == "__main__":
    main()