## Python-codes-to-communicate-at-Gemini-models-via-vertex-AI

## Explanation about the repository and how use the code
  Repository where I have shared a python code that communicate at gemini's models via vertex AI. In this scripts the gemini's models (or also third parts models availables in vertex AI) navigate in       internet and take the most recent and relevant news about a specific topic (for example news about ai, linux, IT, politics and more, to choose the topic You have to edit the prompt in its function       where is explained better below in repository's end) and the code take the model's output, transform the text in json format and write all the N news inside a final word document.
  To execute the a script, You have to configure the first parametres so location (The google's servers geographical area), project ID (Your project's id in google cloud), num of items (how many news      you want in the final document) and the access at vertex ai is ADC this part to edit is explained better below in the repository's end. Before to execute you have to install local gcloud (google cloud   cli) and set default credentials and default project using these 2 commands (these commands are in bash so work mainly in linux terminals):
  # command to execute for set default credentials for ADC access:
    gcloud auth application-default login
  # Command to set the default project (GCP it means google cloud platform):
    gcloud config set project YOUR_GCP_PROJECT_ID

## Parametres to edit inside the code to edit before to use it:

  #-----------------------------------------------------------------------------
  #Configuration
  #-----------------------------------------------------------------------------
  #The lines above are a "block" comment: they only serve to visually
  #separate the logical sections of the file for whoever reads the code; they have
  #no effect on the execution of the program.

  TOPIC = ""  # Insert here the topic that You want the news. for example AI, Politics, Economy, Tech and more else
  #Is a variable that contains a string where the follow string is the topic that You want the news
  #in the final word Document, so insert in the "" the new's topic.

  NEWS_FEATURES = """"""
  #This variable contain a docstring with the "features" about the news for example 
  #if i want all the news about AI, I will insert in the docstring what kind of news I want about ai, so 
  #new updates and models released by AI companies, AI regulaments by governments, futures updates/models released by AI companies and their features,
  #hardware and more else. in this docsting You have to specific what kind of news you want obtain about the topic   
  PROJECT_ID = "your-gcp-project-id"
  #Defines a global variable (constant by convention, written in
  #UPPERCASE) containing the identifier of the Google Cloud project to use
  #to authenticate and bill calls to Vertex AI. This value must
  #correspond to a real project on which the ADC login
  #(Application Default Credentials) described in the initial docstring has been performed.

  #"global" works for the newest 2.5 models; "us-central1" is a safe alternative.
  #Explanatory comment: clarifies the possible alternatives for the
  #next variable (LOCATION), explaining in which cases it is convenient to use one value or
  #the other. It is not executed code.

  LOCATION = ""
  #Defines the geographic/infrastructural region of Google Cloud
  #(data center) to which requests to the Vertex AI model will be
  #routed. Some models or features are only available in certain
  #regions, so this value must be compatible with the models listed below.
  #For example "global", "us-central1" and more else.

  #Separator blank line.

  #Tried in order — automatic failover if one is unavailable / quota-limited.
  #Comment: explains the logic behind the MODELS list defined right below,
  #i.e. that they will be tried in sequence until one responds correctly.

  MODELS = ["gemini-2.5-pro", "gemini-2.5-flash"]
  #Creates a list (ordered, mutable data structure) of strings, each
  #the name of a Gemini model available on Vertex AI. Order matters: the
  #program will first try "gemini-2.5-pro" (more capable but possibly
  #slower/more expensive or subject to quota limits) and, only if it fails, will move to the
  #next one "gemini-2.5-flash" (faster/cheaper), thus implementing a
  #"failover" mechanism (automatic fallback in case of error).

  #Separator blank line.

  NUM_ITEMS = 15            # how many news items to gather Inserted 15 as default but You can edit it and 
  #Defines an integer constant representing how many distinct news items the
  #model must return in total. The comment alongside (after the #)
  #clarifies the meaning of the value; the value itself will then be inserted
  #dynamically into the prompt text sent to the model (see build_prompt).
  
  INCLUDE_SOURCES = True    # append a "Sources" section (the grounded web links)
  #Defines a boolean constant (True/False) that acts as a
  #switch ("feature flag"): if True, a section with the list of
  #web links used by the model as sources during the
  #search ("grounding", i.e. anchoring responses to real sources) will also be added to the final Word document.

 OUTPUT_FILE = f"File_name{datetime.now():%Y-%m-%d}.docx"   # Substitute "File_name" with the name that you want save the final file 
  #Builds the output file path/name using an "f-string"
  #(formatted string): everything between curly braces {} is
  #evaluated as a Python expression and the result is inserted in its
  #place in the final text. Here "datetime.now()" obtains the current
  #date and time, and ":%Y-%m-%d" is a formatting specifier that turns it# into a string like "2026-07-03". The final result is therefore a
  #relative path like "./File_name_2026-07-03.docx", in the
  #current folder from which the script is launched.
# where see here on github code's output files examples
I added in this repository the subdirectory code_outputs_files, where inside that You will find some examples of word's files that code returns as output with the news. 
