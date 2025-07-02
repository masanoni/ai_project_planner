import { GoogleGenAI, GenerateContentResponse } from "@google/genai";
import { ProjectTask, SubStep, SlideDeck, ActionItem, TextboxElement, Slide, ProjectHealthReport, GanttItem } from '../types';
import { GEMINI_MODEL_TEXT } from '../constants';

if (!process.env.API_KEY) {
  throw new Error("API_KEY environment variable not set. Please ensure it is configured.");
}

const ai = new GoogleGenAI({ apiKey: process.env.API_KEY });

/**
 * Sanitizes "dependencies" arrays in a JSON string, which are prone to AI hallucinations.
 * It extracts all valid string literals from the array and reconstructs it.
 * @param jsonStr The raw JSON string.
 * @returns A sanitized JSON string.
 */
const sanitizeDependencyArrays = (jsonStr: string): string => {
  // This regex finds all "dependencies" arrays and captures their content.
  // The 'g' flag ensures all occurrences are replaced.
  const dependenciesRegex = /"dependencies"\s*:\s*\[([\s\S]*?)\]/g;

  return jsonStr.replace(dependenciesRegex, (match, content) => {
    // Inside the array content, find all valid string literals ("...").
    const stringLiterals = content.match(/"[^"]*"/g);
    
    // Reconstruct the array content by joining the found literals with commas.
    // If no valid literals are found, create an empty array.
    const sanitizedContent = stringLiterals ? stringLiterals.join(', ') : '';
    
    // Return the reconstructed "dependencies" array.
    return `"dependencies": [${sanitizedContent}]`;
  });
};


/**
 * Parses a JSON object from a string, attempting to clean it up first.
 * @param text The raw string response from the AI.
 * @returns A parsed object of type T, or null if parsing fails.
 */
const parseJsonFromText = <T,>(text: string): T | null => {
  let jsonStr = text.trim();
  
  // A targeted fix for an observed AI hallucination where it replaces a closing brace and comma with a stray character.
  // This is brittle but can fix specific recurring generation errors.
  // The regex looks for the stray character '棟' between a valid JSON value-ending character and a new object opening brace.
  jsonStr = jsonStr.replace(/(?<=}|"|\d|true|false|null)\s*棟\s*(?=\{)/g, `},`);

  // Sanitize "dependencies" arrays which are prone to string-related hallucinations.
  jsonStr = sanitizeDependencyArrays(jsonStr);

  const fenceRegex = /^```(?:json)?\s*\n?(.*?)\n?\s*```$/s;
  const match = jsonStr.match(fenceRegex);
  if (match && match[1]) {
    jsonStr = match[1].trim();
  } else {
    const firstBrace = jsonStr.indexOf('{');
    const firstBracket = jsonStr.indexOf('[');
    
    if (firstBrace === -1 && firstBracket === -1) {
      console.error("No JSON object or array found in the response string.", text);
      throw new Error(`AI response did not contain a valid JSON structure. Raw response: ${text.substring(0, 500)}...`);
    }

    const start = firstBrace === -1 ? firstBracket : (firstBracket === -1 ? firstBrace : Math.min(firstBrace, firstBracket));
    const lastBrace = jsonStr.lastIndexOf('}');
    const lastBracket = jsonStr.lastIndexOf(']');
    const end = Math.max(lastBrace, lastBracket);
    
    if (start > -1 && end > start) {
      jsonStr = jsonStr.substring(start, end + 1);
    }
  }

  try {
    return JSON.parse(jsonStr) as T;
  } catch (e) {
    console.error("Failed to parse JSON response:", e, "Original text:", text, "Processed string:", jsonStr);
    throw new Error(`Failed to parse AI response as JSON. Raw response: ${text.substring(0, 500)}...`);
  }
};

/**
 * Centralized error handler for Gemini API calls.
 * @param error The error object caught from the API call.
 * @param context A string describing the context of the call (e.g., 'project plan generation').
 */
const handleGeminiError = (error: unknown, context: string): never => {
    console.error(`Error in Gemini API call during ${context}:`, error);
    
    let finalMessage = `AIとの通信中に不明なエラーが発生しました (${context})。`;
    let errorDetails: any = null;

    if (error instanceof Error) {
        finalMessage = `AIとの通信に失敗しました (${context}): ${error.message}`;
        // The error message might contain the JSON
        try {
            errorDetails = JSON.parse(error.message);
        } catch (e) { /* Not a JSON message */ }
    } else if (typeof error === 'object' && error !== null) {
        // The error itself might be the object, e.g. { error: { ... } }
        errorDetails = error;
    }

    // Now check the details we found for specific API error structures
    if (errorDetails) {
        const apiError = errorDetails.error || errorDetails; // The actual error can be nested or be the object itself
        if (apiError && (apiError.status === 'RESOURCE_EXHAUSTED' || apiError.code === 429)) {
            finalMessage = "API利用上限に達しました。Google AI Platformのプランと請求情報を確認してください。(You have exceeded your API quota. Please check your plan and billing details on the Google AI Platform.)";
        } else if (apiError && apiError.message && typeof apiError.message === 'string') {
            finalMessage = `AI APIエラー: ${apiError.message}`;
        }
    }

    throw new Error(finalMessage);
};


export const generateProjectPlan = async (goal: string, date: string): Promise<ProjectTask[]> => {
  const prompt = `
    You are an expert project planner. Your task is to break down a high-level project goal into a sequence of actionable tasks.
    CONTEXT:
    - Project Goal: "${goal}"
    - Target Completion Date: "${date}"
    INSTRUCTIONS:
    1.  Generate a sequence of 3 to 7 high-level, actionable tasks.
    2.  Provide a concise title (max 10 words) and a brief description (1-2 sentences) for each.
    3.  Your response MUST be a single, valid JSON array of objects. Do NOT include any explanations or markdown fences.
    4.  The language of the output should match the language of the input goal.
    5.  Each object MUST have this structure: { "id": "unique_string_id", "title": "Task Title", "description": "Task Description" }
  `;

  try {
    const response: GenerateContentResponse = await ai.models.generateContent({
      model: GEMINI_MODEL_TEXT, contents: prompt, config: { responseMimeType: "application/json" },
    });
    const tasks = parseJsonFromText<ProjectTask[]>(response.text);
    if (!tasks || !Array.isArray(tasks) || tasks.some(task => !task.id || !task.title || !task.description)) {
        throw new Error("AI returned an unexpected format for project tasks.");
    }
    return tasks;
  } catch (error) {
    handleGeminiError(error, 'project plan generation');
  }
};

export const generateStepProposals = async (task: ProjectTask): Promise<{ title: string; description: string; }[]> => {
  const prompt = `
    You are a project management expert. Analyze the given task and propose a list of concrete next steps to accomplish it.
    CONTEXT:
    - Task Title: "${task.title}"
    - Task Description: "${task.description}"
    - Required Resources: "${task.extendedDetails?.resources || 'Not specified'}"
    INSTRUCTIONS:
    1.  Generate a list of 3 to 6 actionable proposals (sub-steps).
    2.  Each proposal must have a concise 'title' and a 'description'.
    3.  Focus on breaking down the main task into logical phases or components.
    4.  Your response MUST be a single, valid JSON array of objects. Do not include any explanations.
    5.  The language must match the input task language.
    6.  The JSON structure for each item must be: { "title": "...", "description": "..." }
  `;
  try {
    const response: GenerateContentResponse = await ai.models.generateContent({
      model: GEMINI_MODEL_TEXT, contents: prompt, config: { responseMimeType: "application/json" },
    });
    const proposals = parseJsonFromText<{ title: string; description: string; }[]>(response.text);
    if (!proposals || !Array.isArray(proposals) || proposals.some(p => !p.title || !p.description)) {
        throw new Error("AI returned an unexpected format for step proposals.");
    }
    if (proposals.length === 0) {
        throw new Error("AI did not generate any proposals. The task might be too abstract.");
    }
    return proposals;
  } catch (error) {
    handleGeminiError(error, 'step proposal generation');
  }
};


export const generateInitialSlideDeck = async (task: ProjectTask, projectGoal: string): Promise<SlideDeck> => {
    const prompt = `
        You are a professional presentation designer and project analyst. Your task is to create a project status report slide deck based on ALL the provided data.
        CONTEXT:
        - Overall Project Goal: "${projectGoal}"
        - Full Task Data (JSON): ${JSON.stringify(task)}
        INSTRUCTIONS:
        1.  Generate a comprehensive slide deck (5-8 slides).
        2.  The response MUST be a single, valid JSON object. Do NOT use markdown.
        3.  The JSON must follow this structure: { "slides": [ { "id": "...", "layout": "...", "isLocked": false, "elements": [ ... ] } ] }.
        4.  Available element types: 'textbox', 'image', 'table', 'chart', 'flowchart'.
        5.  For 'textbox' elements, you MUST use a "content" field for the text.
        6.  **CRITICAL SYNTHESIS**:
            - Create a title slide and an overview slide.
            - **Flowchart Slide**: If the task has sub-steps, create one slide dedicated to visualizing the workflow. This slide should have a 'textbox' title (e.g., "サブステップのワークフロー") and a 'flowchart' element. The 'flowchart' element must have this structure: \`{"id": "...", "type": "flowchart", "position": {"x": 5, "y": 15, "width": 90, "height": 75}, "data": {"subSteps": [...]}}\`. You MUST copy the entire 'subSteps' array from the input JSON into the \`data.subSteps\` field.
            - **ACTION ITEM FOCUS**: For significant Sub-Steps, create additional summary slides. Analyze 'actionItems' and report on completion status.
            - Create slides for key challenges and a final summary.
        7.  Write concise, professional text in the same language as the input task.
        8.  Position elements logically. Do not let them overlap.
    `;
    try {
        const response: GenerateContentResponse = await ai.models.generateContent({
            model: GEMINI_MODEL_TEXT, contents: prompt, config: { responseMimeType: "application/json" },
        });
        const deck = parseJsonFromText<SlideDeck>(response.text);
        if (!deck || !Array.isArray(deck.slides)) {
            throw new Error("AI returned an invalid format for the slide deck.");
        }
        // Validate that textboxes use 'content', not 'text'
        deck.slides.forEach(slide => {
          slide.elements.forEach(el => {
            if (el.type === 'textbox' && typeof (el as any).text !== 'undefined') {
              (el as TextboxElement).content = (el as any).text;
              delete (el as any).text;
            }
          })
        });
        return deck;
    } catch (error) {
        handleGeminiError(error, 'initial slide deck generation');
    }
};

export const regenerateSlideDeck = async (existingDeck: SlideDeck, task: ProjectTask, projectGoal: string): Promise<SlideDeck> => {
    const prompt = `
      You are a presentation designer and project analyst. Your task is to update a project status report slide deck based on new data, while preserving slides that have been manually locked by the user.
      CONTEXT:
      - Overall Project Goal: "${projectGoal}"
      - Updated Full Task Data (JSON): ${JSON.stringify(task)}
      - Existing Slide Deck (with locked slides noted by "isLocked": true): ${JSON.stringify(existingDeck)}

      INSTRUCTIONS:
      1.  Analyze 'existingDeck'. Slides with '"isLocked": true' MUST NOT be changed. Return them exactly as they are.
      2.  For all other slides (unlocked), REGENERATE their content from scratch based on the 'updatedTaskData'.
      3.  This includes regenerating 'flowchart' slides to reflect the latest sub-step data.
      4.  **ACTION ITEM FOCUS**: When regenerating, pay close attention to the progress and reports within each sub-step's 'actionItems'.
      5.  The final output MUST be a single, valid JSON object representing the complete slide deck, with all slides (locked and regenerated) in their original order.
      6.  Follow the same JSON structure and rules as the initial generation.
    `;
    try {
        const response: GenerateContentResponse = await ai.models.generateContent({
            model: GEMINI_MODEL_TEXT, contents: prompt, config: { responseMimeType: "application/json" },
        });
        const deck = parseJsonFromText<SlideDeck>(response.text);
        if (!deck || !Array.isArray(deck.slides)) {
            throw new Error("AI returned an invalid format for the regenerated slide deck.");
        }
        return deck;
    } catch (error) {
        handleGeminiError(error, 'slide deck regeneration');
    }
};


export const optimizeSlideLayout = async (deck: SlideDeck): Promise<SlideDeck> => {
    const prompt = `
        You are an expert presentation designer. The following JSON represents a slide deck.
        Analyze its content and structure. Your task is to improve the layout, wording, and visual hierarchy for maximum clarity and impact.
        You can rephrase text for conciseness, change element positions and sizes, and adjust font properties. Do not change element types or IDs.
        Return the updated slide deck as a valid JSON object with the exact same structure as the input. Do not use markdown.
        CRITICAL: Textbox elements use a "content" field, NOT a "text" field.
        
        Input Deck:
        ${JSON.stringify(deck)}
    `;
    try {
        const response: GenerateContentResponse = await ai.models.generateContent({
            model: GEMINI_MODEL_TEXT, contents: prompt, config: { responseMimeType: "application/json" },
        });
        const optimizedDeck = parseJsonFromText<SlideDeck>(response.text);
        if (!optimizedDeck || !Array.isArray(optimizedDeck.slides)) {
            throw new Error("AI returned an invalid format for the optimized slide deck.");
        }
        return optimizedDeck;
    } catch (error) {
        handleGeminiError(error, 'slide layout optimization');
    }
};

export const generateProjectHealthReport = async (tasks: ProjectTask[], projectGoal: string, targetDate: string): Promise<ProjectHealthReport> => {
    const prompt = `
      You are a senior project manager AI. Your task is to conduct a holistic health check of the entire project.
      CONTEXT:
      - Overall Project Goal: "${projectGoal}"
      - Final Target Date: "${targetDate}"
      - Current Date: "${new Date().toISOString().split('T')[0]}"
      - Full Project Data (Tasks, Sub-steps, Action Items, statuses, due dates): ${JSON.stringify(tasks, null, 2)}

      INSTRUCTIONS:
      1.  **Holistic Analysis**: Review ALL provided data. Compare task/sub-step due dates with the current date. Analyze dependencies, blockers, and the completion rate of action items.
      2.  **Determine Overall Status**: Categorize the project's health as 'On Track', 'At Risk', or 'Off Track'.
      3.  **Identify Positives**: List 2-3 key accomplishments or areas that are progressing well.
      4.  **Identify Concerns**: List the most critical risks or issues. For each, explain WHY it's a concern (e.g., "Task 'X' is 2 weeks overdue and blocking 3 other tasks"). Note the related task IDs.
      5.  **Propose Solutions**: For each major concern, provide concrete, actionable suggestions for improvement. (e.g., "Re-allocate resources from Task Y to Task X", "Hold a risk mitigation meeting for Z").
      6.  **Summarize**: Write a concise, executive-level summary of the project's current state.
      7.  **JSON Output**: Your response MUST be a single, valid JSON object following the ProjectHealthReport structure. Do not include markdown or explanations.
      8.  **The output language MUST be Japanese.**

      JSON Structure to follow:
      {
        "overallStatus": "'On Track' | 'At Risk' | 'Off Track'",
        "summary": "string",
        "positivePoints": ["string"],
        "areasOfConcern": [{ "description": "string", "relatedTaskIds": ["string"] }],
        "suggestions": ["string"]
      }
    `;

    try {
        const response: GenerateContentResponse = await ai.models.generateContent({
            model: GEMINI_MODEL_TEXT, contents: prompt, config: { responseMimeType: "application/json" },
        });
        const report = parseJsonFromText<ProjectHealthReport>(response.text);
        if (!report || !report.overallStatus || !report.summary) {
            throw new Error("AI returned an invalid format for the project health report.");
        }
        return report;
    } catch (error) {
        handleGeminiError(error, 'project health report generation');
    }
};

export const generateProjectReportDeck = async (tasks: ProjectTask[], projectGoal: string, targetDate: string): Promise<SlideDeck> => {
    const prompt = `
        You are a senior project analyst AI. Your task is to create a comprehensive slide deck summarizing the ENTIRE project status.
        CONTEXT:
        - Overall Project Goal: "${projectGoal}"
        - Final Target Date: "${targetDate}"
        - Full Project Data (JSON): ${JSON.stringify(tasks, null, 2)}
        INSTRUCTIONS:
        1.  Synthesize a comprehensive slide deck (6-10 slides) summarizing the entire project.
        2.  The response MUST be a single, valid JSON object following the SlideDeck structure. Do NOT use markdown.
        3.  The output language MUST be Japanese.
        4.  **CRITICAL SYNTHESIS**:
            - **Title & Overview**: Create a title slide and a project overview slide.
            - **Task Summary & Flowchart**: For each major task in the 'tasks' array, create one or two summary slides. One slide should show a \`flowchart\` of its sub-steps if they exist, following the structure \`{"type": "flowchart", "data": {"subSteps": [...]}}\`. You MUST copy the task's specific \`subSteps\` array into the \`data.subSteps\` field. Another slide can summarize the task's status and outcomes.
            - **Key Achievements & Risks**: Create dedicated slides for significant achievements and project-level risks.
            - **Conclusion**: A final slide summarizing the project's outlook and next steps.
        5.  Use the standard JSON format for slides and elements. Remember "content" for textboxes.
    `;
    try {
        const response: GenerateContentResponse = await ai.models.generateContent({
            model: GEMINI_MODEL_TEXT, contents: prompt, config: { responseMimeType: "application/json" },
        });
        const deck = parseJsonFromText<SlideDeck>(response.text);
        if (!deck || !Array.isArray(deck.slides)) {
            throw new Error("AI returned an invalid format for the project slide deck.");
        }
        return deck;
    } catch (error) {
        handleGeminiError(error, 'project report deck generation');
    }
};


export const regenerateProjectReportDeck = async (existingDeck: SlideDeck, tasks: ProjectTask[], projectGoal: string, targetDate: string): Promise<SlideDeck> => {
    const prompt = `
      You are a senior project analyst AI. Your task is to update a project-wide status report slide deck based on new data, while preserving slides that have been manually locked by the user.
      CONTEXT:
      - Overall Project Goal: "${projectGoal}"
      - Final Target Date: "${targetDate}"
      - Updated Full Project Data (JSON): ${JSON.stringify(tasks, null, 2)}
      - Existing Slide Deck (with locked slides noted by "isLocked": true): ${JSON.stringify(existingDeck)}

      INSTRUCTIONS:
      1.  Analyze the 'existingDeck'. Slides with '"isLocked": true' MUST NOT be changed. Return them exactly as they are.
      2.  For all other slides (unlocked slides), REGENERATE their content from scratch based on the updated project data.
      3.  This includes regenerating any 'flowchart' elements on unlocked slides to show the latest sub-step data.
      4.  Synthesize information across all tasks to provide a holistic project view in the regenerated slides.
      5.  The final output MUST be a single, valid JSON object for the complete slide deck, in the original slide order.
      6.  The output language MUST be Japanese.
    `;
    try {
        const response: GenerateContentResponse = await ai.models.generateContent({
            model: GEMINI_MODEL_TEXT, contents: prompt, config: { responseMimeType: "application/json" },
        });
        const deck = parseJsonFromText<SlideDeck>(response.text);
        if (!deck || !Array.isArray(deck.slides)) {
            throw new Error("AI returned an invalid format for the regenerated project slide deck.");
        }
        return deck;
    } catch (error) {
        handleGeminiError(error, 'project report deck regeneration');
    }
};

export const generateGanttData = async (tasks: ProjectTask[], projectGoal: string, targetDate: string): Promise<GanttItem[]> => {
    const prompt = `
      You are a project management assistant AI. Your task is to convert a project structure into data for a Gantt chart.
      CONTEXT:
      - Project Goal: "${projectGoal}"
      - Project Start Date: "${new Date().toISOString().split('T')[0]}"
      - Project Target End Date: "${targetDate}"
      - Full Project Data (JSON): ${JSON.stringify(tasks, null, 2)}

      INSTRUCTIONS:
      1.  Create a flat list of items for the Gantt chart. Include every ProjectTask, every SubStep, and every ActionItem from all tasks.
      2.  For each item, generate a JSON object with this exact structure:
          { "id": "string", "name": "string", "start": "YYYY-MM-DD", "end": "YYYY-MM-DD", "progress": number, "dependencies": ["string"], "type": "'task' | 'substep' | 'actionitem'", "parentId": "string | null" }
      3.  **Date Estimation**:
          - The project must fit between the start and target dates.
          - Use an item's \`dueDate\` as its 'end' date if available for tasks and sub-steps.
          - Estimate 'start' and 'end' dates logically. An item must start after its dependencies end.
          - For 'actionitem', estimate a short duration (e.g., 1-2 days) within the parent sub-step's timeframe.
      4.  **Progress Calculation**:
          - For a 'task': 'Completed' is 100, 'Not Started' is 0. For 'In Progress' or 'Blocked', average the progress of its sub-steps. If no sub-steps, 'In Progress' is 50.
          - For a 'substep': Calculate progress from its \`actionItems\` (% of completed items). If no action items, use its \`status\`: 'Completed' is 100, 'In Progress' is 50, 'Not Started' is 0.
          - For an 'actionitem': 'completed: true' is 100, 'completed: false' is 0.
      5.  **Dependencies & Parent ID**:
          - **Rule**: The 'dependencies' array models the sequential workflow. A dependency can ONLY exist between items of the SAME \`type\` AND with the SAME \`parentId\`.
          - **Tasks**: A task's 'parentId' is ALWAYS null. Use \`nextTaskIds\` to establish dependencies between tasks.
          - **Sub-steps**: A sub-step's 'parentId' MUST BE the ID of its parent ProjectTask. Use \`nextSubStepIds\` to establish dependencies ONLY between sub-steps within the same task.
          - **Action Items**: An action item's 'parentId' MUST BE the ID of its parent SubStep. Create sequential dependencies ONLY between action items within the same sub-step (item 2 depends on item 1, etc.).
          - **CRITICAL**: Do NOT create dependencies across different parents (e.g., from a sub-step in Task A to one in Task B) or across different types (e.g., from a task to a sub-step).
      6.  Your response MUST be a single, valid JSON array of GanttItem objects. Do not use markdown.
    `;

    try {
        const response: GenerateContentResponse = await ai.models.generateContent({
            model: GEMINI_MODEL_TEXT, contents: prompt, config: { responseMimeType: "application/json" },
        });
        const ganttData = parseJsonFromText<GanttItem[]>(response.text);
        if (!ganttData || !Array.isArray(ganttData)) {
            throw new Error("AI returned an invalid format for the Gantt chart data.");
        }
        return ganttData;
    } catch (error) {
        handleGeminiError(error, 'Gantt chart data generation');
    }
};