import os
import torch
import tkinter as tk
from tika import parser
from tkinter import simpledialog
from transformers import DPRQuestionEncoder, DPRQuestionEncoderTokenizer, T5ForConditionalGeneration, T5Tokenizer, DPRContextEncoder, DPRContextEncoderTokenizer



# Preprocess document data
document_titles = []
document_passages = []
count = 1
print('Extracting Text...')
for root,dirs,files in os.walk(r"\\softtest\Dev\ZenQMS\SOPs\subset"):
    for names in files:
        text = parser.from_file(os.path.join(root,names))
        document_titles.append(str(text['content'].replace('\n','').replace('SOP',', SOP').split('Version')[0]).rstrip())
        document_passages.append(str(text['content'].replace('\n','')))
        print(count)
        count += 1

# Initialize DPR question encoder and tokenizer
question_encoder_name = "facebook/dpr-question_encoder-single-nq-base"
question_tokenizer = DPRQuestionEncoderTokenizer.from_pretrained(question_encoder_name)
question_encoder = DPRQuestionEncoder.from_pretrained(question_encoder_name)

# Initialize DPR context encoder and tokenizer
context_encoder_name = "facebook/dpr-ctx_encoder-single-nq-base"
context_tokenizer = DPRContextEncoderTokenizer.from_pretrained(context_encoder_name)
context_encoder = DPRContextEncoder.from_pretrained(context_encoder_name)

# Initialize T5 model and tokenizer
t5_model_name = "google/flan-t5-xxl"
t5_tokenizer = T5Tokenizer.from_pretrained(t5_model_name)
t5_model = T5ForConditionalGeneration.from_pretrained(t5_model_name)

# Example query
root = tk.Tk()
query = tk.simpledialog.askstring('\n\n\n\n\n','\n\n\n\n\n                                  Ask a Question                              \n\n\n\n\n')
root.lift()
root.withdraw()
print('Generating Answer...\n\n\n')

# Encode the query using DPR question encoder
encoded_query = question_tokenizer(query, return_tensors="pt")
question_embedding = question_encoder(**encoded_query).pooler_output

# Encode the document passages using DPR context encoder
encoded_passages = context_tokenizer(document_passages, return_tensors="pt", padding=True, truncation=True)
passage_embeddings = context_encoder(**encoded_passages).pooler_output

# Calculate similarity scores between query and passages
similarities = torch.matmul(question_embedding, passage_embeddings.T)
sorted_indices = torch.argsort(similarities, descending=True)

# Get top k passages
k = 3
top_k_indices = sorted_indices[0, :k]

# Format inputs for T5
formatted_inputs = [f"question: {query} context: {document_passages[i]}" for i in top_k_indices]

# Generate answers using T5
encoded_inputs = t5_tokenizer(formatted_inputs, return_tensors="pt", padding=True)
generated_answers = t5_model.generate(**encoded_inputs)

# Decode generated answers
decoded_answers = [t5_tokenizer.decode(answer, skip_special_tokens=True) for answer in generated_answers]

print("Answers with Similarity Scores and Associated Document Titles:\n")
for i, answer in enumerate(decoded_answers):
    similarity_score = similarities[0, top_k_indices[i]].item()
    document_title = document_titles[top_k_indices[i]]
    print(f"Answer: {answer}, \nDocument Title: {document_title}, \nSimilarity Score: {similarity_score:.4f}\n\n")
