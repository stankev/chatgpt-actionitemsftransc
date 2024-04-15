# chatgpt-actionitemsftransc
Microsoft Word macro to call OpenAI API (chat/completions) to create action items from Audio Transcription

Example of how to use the power of Generative AI from Microsoft Word & OpenAI APIs to create a list of action items.  The input to the process is an audio recording that is first transcribed and then action items generated.

This was tested with the “Microsoft® Word for Microsoft 365 MSO (Version 2402)”.  Other versions of Micosoft Word should work but testing will be needed and possibly slight modifications will the required.

## Instructions
Detailed instructions can be found in this blog post: [How to Efficiently Transcribe Audio Recordings and Generate Action Items with Word and OpenAI APIs](https://www.aiinnovate360.com/post/how-to-efficiently-transcribe-audio-recordings-and-generate-action-items-with-word-and-openai)
## Dependencies
This code has a dependency on two other VBA libraries:
- Requires importing the [VBA-JSON library](https://github.com/VBA-tools/VBA-JSON) providing the JsonConverter
- Requires importing the [VBA-Dictionary](https://github.com/VBA-tools/VBA-Dictionary)
- Requires an [OpenAI API Account](https://openai.com/product) and the API Key

## Disclaimer
The authors and contributors of this program provide it as-is, without any warranties or guarantees. They cannot be held responsible for any damages resulting from the use of this program.

## License
This program is licensed under the MIT license.

## Author
Mark Stankevicius
