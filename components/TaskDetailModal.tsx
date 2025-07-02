
import React, { useEffect, useState, ChangeEvent, useRef, useCallback, useMemo } from 'react';
import { 
  ProjectTask, SubStep, EditableExtendedTaskDetails, TaskStatus,
  SubStepStatus, Attachment, SlideDeck, ActionItem, ActionItemReport
} from '../types';
import LoadingSpinner from './LoadingSpinner';
import ErrorMessage from './ErrorMessage';
import { 
  XIcon, ListIcon, ClockIcon, NotesIcon, ResourcesIcon, ResponsibleIcon,
  LightBulbIcon, PlusIcon, TrashIcon, SubtaskIcon, RefreshIcon, CheckSquareIcon, SquareIcon, PlusCircleIcon,
  UndoIcon, RedoIcon, TableCellsIcon, ArrowsPointingOutIcon, ArrowsPointingInIcon, PaperClipIcon
} from './icons'; 
import FlowConnector from './FlowConnector';
import { generateStepProposals, generateInitialSlideDeck } from '../services/geminiService';
import SlideEditorView from './SlideEditorView';
import ProposalReviewModal from './ProposalReviewModal';
import ActionItemReportModal from './ActionItemReportModal';
import MatrixEditor from './MatrixEditor';

// --- Helper Input Component ---
const DetailInput: React.FC<{label: string, name: string, value: any, onChange: (e: ChangeEvent<HTMLInputElement | HTMLTextAreaElement | HTMLSelectElement>) => void, type?: string, placeholder?: string, rows?: number, disabled?: boolean, icon?: React.ReactNode, required?: boolean}> = 
  ({label, name, value, onChange, type="text", placeholder, rows, disabled=false, icon, required=false}) => (
  <div className="mb-4">
    <label htmlFor={name} className="block text-sm font-medium text-slate-700 mb-1 flex items-center">
      {icon && <span className="mr-1.5 text-slate-500 w-4 h-4">{icon}</span>}
      {label}
      {required && <span className="text-red-500 ml-1">*</span>}
    </label>
    {type === 'textarea' ? (
      <textarea id={name} name={name} value={value} onChange={onChange} rows={rows || 3} placeholder={placeholder} disabled={disabled}
                className="w-full p-2 border border-slate-300 rounded-md shadow-sm focus:ring-blue-500 focus:border-blue-500 text-sm text-slate-900 bg-white placeholder-slate-400 disabled:bg-slate-100"/>
    ) : (
      <input type={type} id={name} name={name} value={value} onChange={onChange} placeholder={placeholder} disabled={disabled}
             className="w-full p-2 border border-slate-300 rounded-md shadow-sm focus:ring-blue-500 focus:border-blue-500 text-sm text-slate-900 bg-white placeholder-slate-400 disabled:bg-slate-100"/>
    )}
  </div>
);

// --- SubStep Card Component ---
const SubStepCard: React.FC<{ subStep: SubStep; onRemove: () => void; onDragStart: (event: React.DragEvent<HTMLDivElement>, subStepId: string) => void; onClick: () => void; cardRef?: React.RefObject<HTMLDivElement>; isSelected?: boolean; onStartConnection: (subStepId: string, event: React.MouseEvent<HTMLDivElement>) => void; onEndConnection: (subStepId: string) => void; }> = React.memo(({ subStep, onRemove, onDragStart, onClick, cardRef, isSelected, onStartConnection, onEndConnection }) => {
  const getStatusBorder = (status?: SubStepStatus) => {
    switch(status) {
      case SubStepStatus.COMPLETED: return 'border-l-4 border-green-500';
      case SubStepStatus.IN_PROGRESS: return 'border-l-4 border-blue-500';
      default: return 'border-l-4 border-slate-300';
    }
  }

  const handleMouseDownOnHandle = (e: React.MouseEvent<HTMLDivElement>) => {
    e.stopPropagation();
    onStartConnection(subStep.id, e);
  };

  return (
    <div ref={cardRef} draggable onDragStart={(e) => onDragStart(e, subStep.id)} onClick={onClick}
      onMouseUp={(e) => { e.stopPropagation(); onEndConnection(subStep.id); }}
      className={`substep-card absolute p-2.5 rounded-lg shadow-sm cursor-grab active:cursor-grabbing bg-white border w-48 h-auto min-h-[60px] text-xs flex flex-col justify-between hover:shadow-md transition-all duration-200 ${isSelected ? 'ring-2 ring-blue-500 shadow-lg' : 'border-slate-300'} ${getStatusBorder(subStep.status)}`}
      style={{ left: subStep.position?.x || 0, top: subStep.position?.y || 0, touchAction: 'none' }}
    >
      <div className="flex justify-between items-start text-slate-400 mb-1">
        <span className="font-mono text-[10px] select-none">ID: ...{subStep.id.slice(-4)}</span>
         <button onClick={(e) => {e.stopPropagation(); onRemove()}} className="text-red-400 hover:text-red-600 p-0.5" title="サブステップ削除">
          <TrashIcon className="w-3.5 h-3.5"/>
        </button>
      </div>
      <p className={`mb-1 break-words text-slate-800 ${subStep.status === SubStepStatus.COMPLETED ? 'line-through text-slate-500' : ''}`}>{subStep.text || "クリックして編集"}</p>
      
      <div
          onMouseDown={handleMouseDownOnHandle}
          className="absolute right-[-5px] top-1/2 -translate-y-1/2 w-3 h-3 bg-blue-500 rounded-full cursor-crosshair hover:scale-125 transition-transform"
          title="ドラッグして接続"
        />
    </div>
  );
});

// --- Action Items Checklist ---
const ActionItemChecklist: React.FC<{ 
    items: ActionItem[];
    onToggle: (itemId: string) => void;
    onAdd: () => void;
    onUpdateText: (itemId: string, text: string) => void;
    onRemove: (itemId: string) => void;
    onOpenReport: (item: ActionItem) => void;
}> = ({ items, onToggle, onAdd, onUpdateText, onRemove, onOpenReport }) => {
  return (
    <div className="space-y-2 mt-2">
       <h6 className="text-sm font-semibold flex justify-between items-center text-slate-600">アクションアイテム</h6>
       <div className="space-y-1 max-h-48 overflow-y-auto pr-2">
            {items.map(item => (
                <div key={item.id} className="flex items-center group bg-slate-50 p-1.5 rounded-md">
                    <button onClick={() => onToggle(item.id)} className="mr-2 flex-shrink-0">
                        {item.completed ? <CheckSquareIcon className="w-5 h-5 text-green-600"/> : <SquareIcon className="w-5 h-5 text-slate-400"/>}
                    </button>
                    <div className="w-full" onDoubleClick={() => onOpenReport(item)}>
                        <input 
                            type="text"
                            value={item.text}
                            onChange={(e) => onUpdateText(item.id, e.target.value)}
                            placeholder="アクションアイテム名"
                            className={`w-full text-sm bg-transparent outline-none p-1 rounded-sm focus:ring-1 focus:ring-blue-500 focus:bg-white ${item.completed ? 'line-through text-slate-500' : 'text-slate-800'}`}
                        />
                    </div>
                    <button onClick={() => onRemove(item.id)} className="ml-2 text-red-400 opacity-0 group-hover:opacity-100 transition-opacity">
                        <TrashIcon className="w-4 h-4" />
                    </button>
                </div>
            ))}
        </div>
        <button onClick={onAdd} className="mt-1 text-xs text-blue-600 hover:text-blue-800 font-medium flex items-center gap-1">
            <PlusCircleIcon className="w-4 h-4"/>
            アイテムを追加
        </button>
    </div>
  );
};


// --- Main Modal Component ---
interface TaskDetailModalProps {
  task: ProjectTask | null;
  onClose: () => void;
  onUpdateTaskCoreInfo: (taskId: string, details: { title: string; description: string; status: TaskStatus }) => void; 
  onUpdateExtendedDetails: (taskId: string, details: EditableExtendedTaskDetails) => void; 
  generateUniqueId: (prefix: string) => string;
  projectGoal: string;
  targetDate: string;
}

const TaskDetailModal: React.FC<TaskDetailModalProps> = ({ task, onClose, onUpdateTaskCoreInfo, onUpdateExtendedDetails, generateUniqueId, projectGoal, targetDate }) => {
  if (!task) return null;

  // Main state
  const [editableTask, setEditableTask] = useState(task);
  const [initialTask, setInitialTask] = useState<ProjectTask>(task);
  const [focus, setFocus] = useState<'none' | 'info' | 'canvas' | 'details'>('none');
  
  // Modal states
  const [isSlideEditorOpen, setIsSlideEditorOpen] = useState(false);
  const [isProposalModalOpen, setIsProposalModalOpen] = useState(false);
  const [isActionReportModalOpen, setIsActionReportModalOpen] = useState(false);
  const [activeActionItem, setActiveActionItem] = useState<ActionItem | null>(null);

  // History states
  const [localHistory, setLocalHistory] = useState<ProjectTask[]>([]);
  const [localRedoHistory, setLocalRedoHistory] = useState<ProjectTask[]>([]);

  // Loading/Error states
  const [isGeneratingPlan, setIsGeneratingPlan] = useState(false);
  const [planError, setPlanError] = useState<string | null>(null);
  const [isLoadingDeck, setIsLoadingDeck] = useState(false);
  const [proposals, setProposals] = useState<{title: string, description: string}[]>([]);

  // Sub-step canvas refs and logic
  const subStepCanvasRef = useRef<HTMLDivElement>(null);
  const [subStepCardRefs, setSubStepCardRefs] = useState<Map<string, React.RefObject<HTMLDivElement>>>(new Map());
  const draggedSubStepIdRef = useRef<string | null>(null);
  const subStepDragOffsetRef = useRef<{ x: number; y: number }>({ x: 0, y: 0 });
  const [selectedSubStepId, setSelectedSubStepId] = useState<string | null>(null);
  
  // Attachment refs
  const taskAttachmentInputRef = useRef<HTMLInputElement>(null);
  const subStepAttachmentInputRef = useRef<HTMLInputElement>(null);

  // Connection drawing state
  const [connectingState, setConnectingState] = useState<{ fromId: string; fromPos: { x: number; y: number } } | null>(null);
  const [mousePos, setMousePos] = useState({ x: 0, y: 0 });

  const isDirty = useMemo(() => JSON.stringify(initialTask) !== JSON.stringify(editableTask), [initialTask, editableTask]);

  useEffect(() => {
    setEditableTask(task);
    setInitialTask(task);
    setLocalHistory([]);
    setLocalRedoHistory([]);
  }, [task]);

  useEffect(() => {
    const newRefs = new Map<string, React.RefObject<HTMLDivElement>>();
    (editableTask.extendedDetails?.subSteps || []).forEach(ss => {
       if (ss && ss.id) newRefs.set(ss.id, subStepCardRefs.get(ss.id) || React.createRef<HTMLDivElement>());
    });
    setSubStepCardRefs(newRefs);
  }, [editableTask.extendedDetails?.subSteps]);

  const subStepConnectors = useMemo(() => {
      const connectors: Array<{id: string, from: {x:number, y:number}, to: {x:number, y:number}, sourceId: string, targetId: string}> = [];
      (editableTask.extendedDetails?.subSteps || []).forEach(sourceSS => {
        if (sourceSS?.nextSubStepIds?.length) {
          const sourceCardElement = subStepCardRefs.get(sourceSS.id)?.current;
          if (sourceCardElement && sourceSS.position) {
            const sourcePos = { x: sourceSS.position.x + sourceCardElement.offsetWidth, y: sourceSS.position.y + sourceCardElement.offsetHeight / 2 };
            sourceSS.nextSubStepIds.forEach(targetId => {
              const targetSS = editableTask.extendedDetails?.subSteps.find(t => t.id === targetId);
              const targetCardElement = subStepCardRefs.get(targetId)?.current;
              if (targetSS && targetCardElement && targetSS.position) {
                const targetPos = { x: targetSS.position.x, y: targetSS.position.y + targetCardElement.offsetHeight / 2 };
                connectors.push({ id: `subconn-${sourceSS.id}-${targetId}`, from: sourcePos, to: targetPos, sourceId: sourceSS.id, targetId: targetId });
              }
            });
          }
        }
      });
      return connectors;
  }, [editableTask.extendedDetails?.subSteps, subStepCardRefs]);

  const updateEditableTaskWithHistory = useCallback((updater: (task: ProjectTask) => ProjectTask) => {
    setEditableTask(currentTask => {
      setLocalHistory(prev => [...prev, currentTask]);
      setLocalRedoHistory([]);
      return updater(currentTask);
    });
  }, []);

  const updateTask = (updates: Partial<ProjectTask> | ((task: ProjectTask) => Partial<ProjectTask>)) => {
    updateEditableTaskWithHistory(prev => ({...prev, ...(typeof updates === 'function' ? updates(prev) : updates)}));
  };

  const updateExtended = (updates: Partial<EditableExtendedTaskDetails> | ((details: EditableExtendedTaskDetails) => Partial<EditableExtendedTaskDetails>)) => {
    updateTask(prev => ({ extendedDetails: { ...prev.extendedDetails!, ...(typeof updates === 'function' ? updates(prev.extendedDetails!) : updates) }}));
  };

  const handleLocalUndo = () => {
    if (localHistory.length === 0) return;
    const previousState = localHistory[localHistory.length - 1];
    setLocalHistory(prev => prev.slice(0, -1));
    setLocalRedoHistory(prev => [editableTask, ...prev]);
    setEditableTask(previousState);
  };
  
  const handleLocalRedo = () => {
    if (localRedoHistory.length === 0) return;
    const nextState = localRedoHistory[0];
    setLocalRedoHistory(prev => prev.slice(1));
    setLocalHistory(prev => [...prev, editableTask]);
    setEditableTask(nextState);
  };

  const handleInitiateAIPlan = async () => {
    setIsGeneratingPlan(true);
    setPlanError(null);
    try {
      const result = await generateStepProposals(editableTask);
      setProposals(result);
      setIsProposalModalOpen(true);
    } catch (err) {
      setPlanError(err instanceof Error ? err.message : '計画の提案生成に失敗しました。');
    } finally {
      setIsGeneratingPlan(false);
    }
  };

  const handleConfirmProposals = (additions: { newSubSteps: { title: string, description: string }[], newActionItems: { targetSubStepId: string, title: string }[] }) => {
    updateExtended(d => {
        const existingSubSteps = d.subSteps || [];
        const i = existingSubSteps.length;

        const createdSubSteps: SubStep[] = additions.newSubSteps.map((p, index) => ({
            id: generateUniqueId('sub_ai'),
            text: p.title,
            notes: p.description,
            status: SubStepStatus.NOT_STARTED,
            actionItems: [],
            attachments: [],
            position: { x: 10 + ((i + index) % 4) * 210, y: Math.floor((i + index) / 4) * 90 + 10 },
        }));

        const updatedSubSteps = existingSubSteps.map(ss => {
            const itemsToAdd = additions.newActionItems.filter(item => item.targetSubStepId === ss.id);
            if (itemsToAdd.length === 0) return ss;
            const newActionItems: ActionItem[] = itemsToAdd.map(item => ({
                id: generateUniqueId('action_ai'),
                text: item.title,
                completed: false,
            }));
            return {...ss, actionItems: [...(ss.actionItems || []), ...newActionItems]};
        });

        return { subSteps: [...updatedSubSteps, ...createdSubSteps] };
    });
    setIsProposalModalOpen(false);
  };

  const handleAddSubStep = () => {
    const i = editableTask.extendedDetails?.subSteps?.length || 0;
    const newSubStep: SubStep = { id: generateUniqueId('sub_new'), text: "新しいステップ", status: SubStepStatus.NOT_STARTED, actionItems: [], attachments: [], position: { x: 10 + (i % 4) * 210, y: Math.floor(i / 4) * 90 + 10 } };
    updateExtended(d => ({ subSteps: [...(d.subSteps || []), newSubStep] }));
    setSelectedSubStepId(newSubStep.id);
  };
  
  const handleUpdateSubStep = useCallback((subStepId: string, updates: Partial<SubStep> | ((ss: SubStep) => Partial<SubStep>)) => {
     updateExtended(d => ({ subSteps: (d.subSteps || []).map(ss => ss.id === subStepId ? {...ss, ...(typeof updates === 'function' ? updates(ss) : updates)} : ss)}));
  }, [updateExtended]);
  
  const handleRemoveSubStep = useCallback((idToRemove: string) => {
    setSelectedSubStepId(prev => prev === idToRemove ? null : prev);
    updateExtended(d => ({ subSteps: (d.subSteps || []).filter(ss => ss.id !== idToRemove).map(ss => ({ ...ss, nextSubStepIds: ss.nextSubStepIds?.filter(id => id !== idToRemove) }))}));
  }, [updateExtended]);

  const handleAutoLayoutSubSteps = () => {
    const subSteps = editableTask.extendedDetails?.subSteps;
    const canvas = subStepCanvasRef.current;
    if (!subSteps || subSteps.length === 0 || !canvas) return;

    // --- 1. Topological Sort ---
    const adj = new Map<string, string[]>();
    const inDegree = new Map<string, number>();

    subSteps.forEach(ss => {
      adj.set(ss.id, []);
      inDegree.set(ss.id, 0);
    });

    subSteps.forEach(ss => {
      (ss.nextSubStepIds || []).forEach(nextId => {
        if (adj.has(nextId)) {
          adj.get(ss.id)!.push(nextId);
          inDegree.set(nextId, inDegree.get(nextId)! + 1);
        }
      });
    });

    const queue = subSteps.filter(ss => inDegree.get(ss.id) === 0).map(ss => ss.id);
    const levels: string[][] = [];
    let visitedCount = 0;

    while (queue.length > 0) {
      const levelSize = queue.length;
      const currentLevel: string[] = [];
      for (let i = 0; i < levelSize; i++) {
        const u = queue.shift()!;
        currentLevel.push(u);
        visitedCount++;
        (adj.get(u) || []).forEach(v => {
          inDegree.set(v, inDegree.get(v)! - 1);
          if (inDegree.get(v)! === 0) queue.push(v);
        });
      }
      levels.push(currentLevel);
    }

    if (visitedCount < subSteps.length) {
      const laidOutNodes = new Set(levels.flat());
      const remainingNodes = subSteps.filter(ss => !laidOutNodes.has(ss.id)).map(ss => ss.id);
      if (remainingNodes.length > 0) levels.push(remainingNodes);
    }

    // --- 2. New Layout Logic with Wrapping ---
    const canvasWidth = canvas.clientWidth;
    const cardWidth = 192, cardHeight = 76, hSpacing = 60, vSpacing = 20;
    const newPositions = new Map<string, { x: number; y: number }>();
    
    let currentX = 10;
    let currentY = 10;
    let rowMaxHeight = 0;

    levels.forEach((levelNodes) => {
      const columnWidth = cardWidth + hSpacing;
      const columnHeight = levelNodes.length * (cardHeight + vSpacing) - vSpacing;
      
      // Check if this column fits. Use `currentX > 10` to allow the first column to always be placed.
      if (currentX > 10 && currentX + columnWidth > canvasWidth) {
        currentX = 10;
        currentY += rowMaxHeight + vSpacing * 2; // Add extra spacing between visual rows
        rowMaxHeight = 0;
      }
      
      levelNodes.forEach((nodeId, nodeIndexInLevel) => {
        newPositions.set(nodeId, {
          x: currentX,
          y: currentY + nodeIndexInLevel * (cardHeight + vSpacing)
        });
      });

      currentX += columnWidth;
      rowMaxHeight = Math.max(rowMaxHeight, columnHeight);
    });

    // --- 3. Resize Canvas and Update State ---
    let requiredWidth = 0;
    let requiredHeight = 0;
    newPositions.forEach(pos => {
      requiredWidth = Math.max(requiredWidth, pos.x + cardWidth);
      requiredHeight = Math.max(requiredHeight, pos.y + cardHeight);
    });
    
    const innerCanvas = canvas.querySelector('.substep-inner-canvas') as HTMLDivElement;
    if (innerCanvas) {
      innerCanvas.style.width = `${requiredWidth + 20}px`;
      innerCanvas.style.height = `${requiredHeight + 20}px`;
    }
    
    updateExtended(d => ({ 
      subSteps: (d.subSteps || []).map(ss => ({ ...ss, position: newPositions.get(ss.id) || ss.position }))
    }));
  };
  

  const handleSubStepDragStart = (event: React.DragEvent<HTMLDivElement>, subStepId: string) => {
    draggedSubStepIdRef.current = subStepId;
    const cardRect = event.currentTarget.getBoundingClientRect();
    subStepDragOffsetRef.current = {
      x: event.clientX - cardRect.left,
      y: event.clientY - cardRect.top,
    };
    event.dataTransfer.effectAllowed = 'move';
  };

  const handleSubStepDragOver = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    event.dataTransfer.dropEffect = 'move';
  };
  
  const handleSubStepDrop = (event: React.DragEvent<HTMLDivElement>) => {
    event.preventDefault();
    if (!draggedSubStepIdRef.current || !subStepCanvasRef.current) return;
  
    const canvasRect = subStepCanvasRef.current.getBoundingClientRect();
    const draggedCardElement = subStepCardRefs.get(draggedSubStepIdRef.current)?.current;
    if (!draggedCardElement) return;
  
    const cardWidth = draggedCardElement.offsetWidth;
    const cardHeight = draggedCardElement.offsetHeight;
  
    let newX = event.clientX - canvasRect.left - subStepDragOffsetRef.current.x + subStepCanvasRef.current.scrollLeft;
    let newY = event.clientY - canvasRect.top - subStepDragOffsetRef.current.y + subStepCanvasRef.current.scrollTop;
    
    const innerCanvas = subStepCanvasRef.current.querySelector('.substep-inner-canvas') as HTMLDivElement;
    if (innerCanvas) {
       newX = Math.max(0, Math.min(newX, innerCanvas.scrollWidth - cardWidth));
       newY = Math.max(0, Math.min(newY, innerCanvas.scrollHeight - cardHeight));
    }
  
    handleUpdateSubStep(draggedSubStepIdRef.current, { position: { x: newX, y: newY } });
    draggedSubStepIdRef.current = null;
  };

  const handleOpenActionReport = (item: ActionItem) => {
    setActiveActionItem(item);
    setIsActionReportModalOpen(true);
  };

  const handleAttachmentChange = (
    event: ChangeEvent<HTMLInputElement>,
    target: 'task' | 'substep'
  ) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
        const newAttachment: Attachment = {
            id: generateUniqueId('attach'),
            name: file.name,
            type: file.type,
            dataUrl: e.target?.result as string,
        };

        if (target === 'task') {
            updateExtended(d => ({ attachments: [...(d.attachments || []), newAttachment] }));
        } else if (target === 'substep' && selectedSubStepId) {
            handleUpdateSubStep(selectedSubStepId, ss => ({
                attachments: [...(ss.attachments || []), newAttachment]
            }));
        }
    };
    reader.readAsDataURL(file);
    if(event.target) event.target.value = ''; // Reset file input to allow re-uploading the same file
  };
  
  const handleRemoveAttachment = (id: string, target: 'task' | 'substep') => {
    if (target === 'task') {
        updateExtended(d => ({ attachments: d.attachments?.filter(a => a.id !== id) }));
    } else if (target === 'substep' && selectedSubStepId) {
        handleUpdateSubStep(selectedSubStepId, ss => ({
            attachments: ss.attachments?.filter(a => a.id !== id)
        }));
    }
  };
  
  // --- Connection Logic ---

  const handleStartConnection = (fromId: string, event: React.MouseEvent<HTMLDivElement>) => {
    if (!subStepCanvasRef.current) return;
    const canvasRect = subStepCanvasRef.current.getBoundingClientRect();
    const fromPos = {
      x: event.clientX - canvasRect.left + subStepCanvasRef.current.scrollLeft,
      y: event.clientY - canvasRect.top + subStepCanvasRef.current.scrollTop,
    };
    setConnectingState({ fromId, fromPos });
  };
  
  const handleEndConnection = (targetId: string) => {
    if (!connectingState) return;
    const { fromId } = connectingState;
    if (fromId === targetId) { // No self-connections
      setConnectingState(null);
      return;
    }
    // Add connection
    handleUpdateSubStep(fromId, (ss) => ({
      ...ss,
      nextSubStepIds: Array.from(new Set([...(ss.nextSubStepIds || []), targetId]))
    }));
    setConnectingState(null);
  };
  
  const handleDeleteConnection = (sourceId: string, targetId: string) => {
    handleUpdateSubStep(sourceId, ss => ({
      ...ss,
      nextSubStepIds: ss.nextSubStepIds?.filter(id => id !== targetId)
    }));
  };

  const handleCanvasMouseMove = (e: React.MouseEvent<HTMLDivElement>) => {
    if (!connectingState || !subStepCanvasRef.current) return;
    const canvasRect = subStepCanvasRef.current.getBoundingClientRect();
    setMousePos({
      x: e.clientX - canvasRect.left + subStepCanvasRef.current.scrollLeft,
      y: e.clientY - canvasRect.top + subStepCanvasRef.current.scrollTop,
    });
  };

  const handleCanvasMouseUp = () => {
    if (connectingState) setConnectingState(null); // Cancel connection
  };

  // --- End Connection Logic ---


  const handleSaveActionItemReport = (subStepId: string, updatedItem: ActionItem) => {
    handleUpdateSubStep(subStepId, ss => ({ actionItems: ss.actionItems?.map(item => item.id === updatedItem.id ? updatedItem : item) }));
    setIsActionReportModalOpen(false);
    setActiveActionItem(null);
  };
  
  const handleActionItemUpdate = (subStepId: string, itemId: string, updates: Partial<ActionItem>) => {
    handleUpdateSubStep(subStepId, ss => ({ actionItems: ss.actionItems?.map(item => item.id === itemId ? { ...item, ...updates } : item)}));
  };

  const handleToggleActionItem = (subStepId: string, itemId: string) => {
    updateExtended(d => {
        const newSubSteps = (d.subSteps || []).map(ss => {
            if (ss.id !== subStepId) return ss;
            
            const newActionItems = ss.actionItems?.map(item => 
                item.id === itemId ? { ...item, completed: !item.completed } : item
            );
            
            const allCompleted = newActionItems?.every(item => item.completed);
            const someInProgress = newActionItems?.some(item => item.completed);

            let newStatus = ss.status;
            if (newActionItems && newActionItems.length > 0) {
              if (allCompleted) {
                newStatus = SubStepStatus.COMPLETED;
              } else if (someInProgress) {
                newStatus = SubStepStatus.IN_PROGRESS;
              } else {
                newStatus = SubStepStatus.NOT_STARTED;
              }
            } else {
              newStatus = SubStepStatus.NOT_STARTED;
            }

            return { ...ss, actionItems: newActionItems, status: newStatus };
        });
        return { subSteps: newSubSteps };
    });
  };

  const handleAddActionItem = (subStepId: string) => {
    handleUpdateSubStep(subStepId, ss => ({ actionItems: [...(ss.actionItems || []), { id: generateUniqueId('action'), text: '新しいアクション', completed: false }] }));
  };
  const handleRemoveActionItem = (subStepId: string, itemId: string) => {
     handleUpdateSubStep(subStepId, ss => ({ actionItems: ss.actionItems?.filter(item => item.id !== itemId) }));
  };
  
  const handleSaveChanges = () => {
    onUpdateTaskCoreInfo(editableTask.id, { title: editableTask.title, description: editableTask.description, status: editableTask.status! });
    onUpdateExtendedDetails(editableTask.id, editableTask.extendedDetails!);
    onClose(); 
  };
  
  const handleReportDeckSave = (deck: SlideDeck) => {
    updateExtended({ reportDeck: deck });
  };
  
  const handleOpenReportEditor = async () => {
    if (editableTask.extendedDetails?.reportDeck) {
      setIsSlideEditorOpen(true);
      return;
    }
    setIsLoadingDeck(true);
    try {
      const deck = await generateInitialSlideDeck(editableTask, projectGoal);
      updateExtended({ reportDeck: deck });
      setIsSlideEditorOpen(true);
    } catch (error) {
      alert(`レポート生成エラー: ${error instanceof Error ? error.message : '不明なエラー'}`);
    } finally {
      setIsLoadingDeck(false);
    }
  };

  const handleAttemptClose = () => {
    if (!isDirty) {
      onClose();
      return;
    }
    if (window.confirm("変更が保存されていません。現在の変更を破棄してよろしいですか？\n\n・「OK」で変更を破棄して戻る\n・「キャンセル」で編集を続ける（手動で保存できます）")) {
      onClose();
    }
  };

  const selectedSubStep = editableTask.extendedDetails?.subSteps?.find(ss => ss.id === selectedSubStepId) || null;

  // --- RENDER LOGIC ---

  const renderAttachments = (
    attachments: Attachment[] | undefined,
    onRemove: (id: string) => void
  ) => (
    <div className="grid grid-cols-2 sm:grid-cols-3 gap-2 mt-1 max-h-40 overflow-y-auto p-2 border rounded-md bg-slate-50">
      {!attachments || attachments.length === 0 ? (
        <p className="col-span-full text-center text-xs text-slate-400 py-2">添付ファイルはありません</p>
      ) : (
        attachments.map(att => (
          <div key={att.id} className="relative group border rounded-md overflow-hidden bg-white shadow-sm">
            {att.type.startsWith('image/') ? (
              <img src={att.dataUrl} alt={att.name} className="w-full h-20 object-cover" />
            ) : (
              <a href={att.dataUrl} download={att.name} className="w-full h-20 bg-slate-100 flex flex-col items-center justify-center p-1 hover:bg-slate-200" title={`Download ${att.name}`}>
                <PaperClipIcon className="w-5 h-5 text-slate-500" />
              </a>
            )}
            <div className="absolute bottom-0 w-full bg-black bg-opacity-60 p-1">
               <p className="text-white text-[10px] truncate" title={att.name}>{att.name}</p>
            </div>
            <button onClick={() => onRemove(att.id)} className="absolute top-1 right-1 bg-black bg-opacity-70 text-white rounded-full p-0.5 opacity-0 group-hover:opacity-100 transition-opacity">
              <TrashIcon className="w-3 h-3" />
            </button>
          </div>
        ))
      )}
    </div>
  );


  if (isSlideEditorOpen && editableTask.extendedDetails?.reportDeck) {
    return <SlideEditorView 
        tasks={[editableTask]}
        initialDeck={editableTask.extendedDetails.reportDeck} 
        onSave={handleReportDeckSave} 
        onClose={() => setIsSlideEditorOpen(false)} 
        generateUniqueId={generateUniqueId}
        projectGoal={projectGoal}
        targetDate={targetDate}
        reportScope="task"
    />;
  }
  
  if (isProposalModalOpen) {
    return <ProposalReviewModal proposals={proposals} existingSubSteps={editableTask.extendedDetails?.subSteps || []} onConfirm={handleConfirmProposals} onClose={() => setIsProposalModalOpen(false)} />
  }

  if (isActionReportModalOpen && activeActionItem && selectedSubStep) {
    return <ActionItemReportModal 
        actionItem={activeActionItem}
        onSave={(updatedItem) => handleSaveActionItemReport(selectedSubStep.id, updatedItem)}
        onClose={() => setIsActionReportModalOpen(false)}
        generateUniqueId={generateUniqueId}
    />
  }

  return (
    <div className="min-h-screen bg-slate-100 flex flex-col">
      <header className="flex-shrink-0 bg-white shadow-md z-10 sticky top-0">
        <div className="flex items-center justify-between p-4 sm:p-5 w-full max-w-screen-2xl mx-auto">
          <h3 className="text-lg sm:text-xl font-bold text-slate-800 truncate pr-2">{editableTask.title} - 詳細計画</h3>
          <div className="flex items-center space-x-2 sm:space-x-4">
              <button onClick={handleAttemptClose} className="px-3 sm:px-5 py-2 bg-slate-200 text-slate-700 text-sm font-medium rounded-md hover:bg-slate-300">フローに戻る</button>
              <button onClick={handleSaveChanges} className="px-3 sm:px-5 py-2 bg-blue-600 text-white text-sm font-semibold rounded-md hover:bg-blue-700">計画の変更を保存</button>
              <button onClick={handleOpenReportEditor} disabled={isLoadingDeck} className="px-3 sm:px-5 py-2 bg-purple-600 text-white text-sm font-semibold rounded-md hover:bg-purple-700 disabled:bg-slate-400">
                  {isLoadingDeck ? <LoadingSpinner size="sm" /> : 'レポート作成'}
              </button>
          </div>
        </div>
      </header>

      <main className="flex-grow p-4 sm:p-6 overflow-hidden w-full max-w-screen-2xl mx-auto">
        <div className="flex gap-6 h-full">
            { /* Panel 1: Task Info */ }
            <section className={`flex flex-col space-y-4 bg-white p-4 rounded-lg shadow-sm border transition-all duration-300 ease-in-out
                ${focus === 'info' ? 'flex-1' :
                  focus === 'none' ? 'w-[360px] min-w-[360px]' :
                  'w-0 p-0 m-0 border-0 overflow-hidden opacity-0'}`}>
                <div className="flex items-center justify-between border-b pb-2">
                    <h4 className="text-lg font-semibold text-slate-700 flex items-center"><ListIcon className="w-5 h-5 mr-2 text-purple-600"/>タスク情報</h4>
                    <button onClick={() => setFocus(focus === 'info' ? 'none' : 'info')} className="p-1 text-slate-500 hover:text-blue-600" title={focus === 'info' ? '元に戻す' : '最大化'}>
                      {focus === 'info' ? <ArrowsPointingInIcon className="w-5 h-5"/> : <ArrowsPointingOutIcon className="w-5 h-5"/>}
                    </button>
                </div>
                <div className="flex-grow overflow-y-auto pr-2">
                  <DetailInput label="タスクタイトル" name="title" value={editableTask.title} onChange={e => updateTask({title: e.target.value})} required />
                  <DetailInput label="タスク説明" name="description" value={editableTask.description} onChange={e => updateTask({description: e.target.value})} type="textarea" rows={3} required />
                  <DetailInput icon={<ResponsibleIcon />} label="担当者/チーム" name="responsible" value={editableTask.extendedDetails!.responsible} onChange={e => updateExtended({responsible: e.target.value})} />
                  <DetailInput icon={<ClockIcon />} label="このタスクの期日" name="dueDate" type="date" value={editableTask.extendedDetails!.dueDate || ''} onChange={e => updateExtended({dueDate: e.target.value})} />
                  
                  <div>
                    <DetailInput icon={<ResourcesIcon />} label="必要なリソース" name="resources" value={editableTask.extendedDetails!.resources} onChange={e => updateExtended({resources: e.target.value})} type="textarea" />
                    {editableTask.extendedDetails?.resourceMatrix ? (
                      <div className="mt-2">
                        <MatrixEditor
                          matrixData={editableTask.extendedDetails.resourceMatrix}
                          onUpdate={newData => updateExtended({ resourceMatrix: newData })}
                        />
                      </div>
                    ) : (
                      <button onClick={() => updateExtended({ resourceMatrix: { headers: ['購入品', '価格'], rows: [['', '']] } })} className="mt-1 text-xs text-blue-600 hover:text-blue-800 font-medium flex items-center gap-1">
                        <TableCellsIcon className="w-4 h-4" />
                        マトリクスを追加
                      </button>
                    )}
                  </div>

                  <DetailInput icon={<NotesIcon />} label="備考・詳細メモ" name="notes" value={editableTask.extendedDetails!.notes} onChange={e => updateExtended({notes: e.target.value})} type="textarea" rows={4} />
                  
                  <div className="pt-2">
                      <label className="flex items-center justify-between text-sm font-medium text-slate-700 mb-1">
                          タスク全体の添付資料
                          <button onClick={() => taskAttachmentInputRef.current?.click()} className="p-1 hover:bg-slate-100 rounded-full" title="添付ファイルを追加">
                              <PaperClipIcon className="w-4 h-4 text-slate-600"/>
                          </button>
                      </label>
                      {renderAttachments(editableTask.extendedDetails?.attachments, (id) => handleRemoveAttachment(id, 'task'))}
                      <input type="file" ref={taskAttachmentInputRef} onChange={(e) => handleAttachmentChange(e, 'task')} className="hidden"/>
                  </div>
                  
                  <div className="pt-4"><h4 className="text-lg font-semibold text-slate-700 mb-3 border-b pb-2 flex items-center"><LightBulbIcon className="w-5 h-5 mr-2 text-yellow-500"/>AIによる自動計画</h4>
                    <button onClick={handleInitiateAIPlan} disabled={isGeneratingPlan} className="w-full px-4 py-2 text-sm font-medium text-white bg-indigo-600 rounded-md shadow-sm hover:bg-indigo-700 disabled:bg-slate-400">
                      {isGeneratingPlan ? <LoadingSpinner size="sm" /> : 'AIでサブステップを自動計画'}
                    </button>
                    {planError && <div className="mt-2"><ErrorMessage message={planError}/></div>}
                  </div>
                </div>
            </section>
            
            { /* Panel 2: Sub-step Canvas */ }
            <section className={`flex flex-col space-y-4 min-w-0 transition-all duration-300 ease-in-out
                ${focus === 'canvas' ? 'flex-1' :
                  focus === 'none' ? 'flex-1' :
                  'hidden'}`}>
                  <div className="flex items-center justify-between border-b pb-2">
                    <h4 className="text-lg font-semibold text-slate-700 flex items-center"><SubtaskIcon className="w-5 h-5 mr-2 text-blue-600"/>サブステップ計画</h4>
                    <div className="flex gap-2 items-center">
                      <button onClick={handleLocalUndo} disabled={localHistory.length === 0} className="text-xs px-2 py-1 bg-slate-200 text-slate-700 rounded-md hover:bg-slate-300 disabled:opacity-50 disabled:cursor-not-allowed flex items-center gap-1" title="元に戻す"><UndoIcon className="w-3 h-3"/></button>
                      <button onClick={handleLocalRedo} disabled={localRedoHistory.length === 0} className="text-xs px-2 py-1 bg-slate-200 text-slate-700 rounded-md hover:bg-slate-300 disabled:opacity-50 disabled:cursor-not-allowed flex items-center gap-1" title="やり直し"><RedoIcon className="w-3 h-3"/></button>
                      <button onClick={handleAddSubStep} className="text-xs px-2 py-1 bg-blue-500 text-white rounded-md hover:bg-blue-600 flex items-center gap-1"><PlusIcon className="w-3 h-3"/>追加</button>
                      <button onClick={handleAutoLayoutSubSteps} className="text-xs px-2 py-1 bg-slate-200 text-slate-700 rounded-md hover:bg-slate-300 flex items-center gap-1" title="自動整列"><RefreshIcon className="w-3 h-3"/>整列</button>
                      <button onClick={() => setFocus(focus === 'canvas' ? 'none' : 'canvas')} className="p-1 text-slate-500 hover:text-blue-600" title={focus === 'canvas' ? '元に戻す' : '最大化'}>
                          {focus === 'canvas' ? <ArrowsPointingInIcon className="w-5 h-5"/> : <ArrowsPointingOutIcon className="w-5 h-5"/>}
                      </button>
                    </div>
                  </div>
                  <div ref={subStepCanvasRef} onMouseMove={handleCanvasMouseMove} onMouseUp={handleCanvasMouseUp} onDragOver={handleSubStepDragOver} onDrop={handleSubStepDrop}
                       className={`substep-canvas relative border rounded-md bg-white p-2 overflow-auto shadow-inner ${focus === 'canvas' ? '' : 'flex-grow'}`}>
                    <div className="substep-inner-canvas relative"> 
                      {(editableTask.extendedDetails?.subSteps || []).map(ss => (ss ? <SubStepCard key={ss.id} subStep={ss} cardRef={subStepCardRefs.get(ss.id)} isSelected={selectedSubStepId === ss.id} onClick={() => setSelectedSubStepId(ss.id)} onRemove={() => handleRemoveSubStep(ss.id)} onDragStart={handleSubStepDragStart} onStartConnection={handleStartConnection} onEndConnection={handleEndConnection} /> : null))}
                      {subStepConnectors.map(conn => <FlowConnector key={conn.id} from={conn.from} to={conn.to} id={conn.id} onDelete={() => handleDeleteConnection(conn.sourceId, conn.targetId)} />)}
                      {connectingState && <FlowConnector from={connectingState.fromPos} to={mousePos} id="preview-connector" />}
                    </div>
                  </div>
            </section>
            
            { /* Panel 3: Sub-step Details */ }
            <section className={`flex flex-col space-y-4 bg-white p-4 rounded-lg shadow-sm border transition-all duration-300 ease-in-out
                ${focus === 'details' ? 'flex-1' :
                  focus === 'none' ? 'w-[360px] min-w-[360px]' :
                  'hidden'}`}>
                <div className="flex items-center justify-between border-b pb-2">
                  <h5 className="text-lg font-semibold text-slate-800">サブステップ詳細</h5>
                  <button onClick={() => setFocus(focus === 'details' ? 'none' : 'details')} disabled={!selectedSubStep} className="p-1 text-slate-500 hover:text-blue-600 disabled:text-slate-300 disabled:cursor-not-allowed" title={focus === 'details' ? '元に戻す' : '最大化'}>
                    {focus === 'details' ? <ArrowsPointingInIcon className="w-5 h-5"/> : <ArrowsPointingOutIcon className="w-5 h-5"/>}
                  </button>
                </div>
                <div className="flex-grow overflow-y-auto pr-2">
                    {selectedSubStep ? (
                      <div className="space-y-4">
                        <DetailInput label="テキスト" name="text" value={selectedSubStep.text} onChange={(e) => handleUpdateSubStep(selectedSubStep.id, {text: e.target.value})} type="textarea" />
                        <DetailInput label="担当者" name="responsible" value={selectedSubStep.responsible || ''} onChange={(e) => handleUpdateSubStep(selectedSubStep.id, {responsible: e.target.value})} />
                        <DetailInput label="期日" name="dueDate" type="date" value={selectedSubStep.dueDate || ''} onChange={(e) => handleUpdateSubStep(selectedSubStep.id, {dueDate: e.target.value})} />
                        <div><label className="block text-sm font-medium text-slate-700 mb-1">ステータス</label>
                          <select value={selectedSubStep.status || SubStepStatus.NOT_STARTED} onChange={(e) => handleUpdateSubStep(selectedSubStep.id, {status: e.target.value as SubStepStatus})} className="w-full p-2 border rounded-md text-sm bg-white text-slate-900">
                            {Object.values(SubStepStatus).map(s => <option key={s} value={s}>{s}</option>)}</select></div>
                        <DetailInput label="メモ" name="notes" value={selectedSubStep.notes || ''} onChange={(e) => handleUpdateSubStep(selectedSubStep.id, {notes: e.target.value})} type="textarea" />
                         <ActionItemChecklist 
                            items={selectedSubStep.actionItems || []}
                            onToggle={(itemId) => handleToggleActionItem(selectedSubStep.id, itemId)}
                            onAdd={() => handleAddActionItem(selectedSubStep.id)}
                            onUpdateText={(itemId, text) => handleActionItemUpdate(selectedSubStep.id, itemId, { text })}
                            onRemove={(itemId) => handleRemoveActionItem(selectedSubStep.id, itemId)}
                            onOpenReport={(item) => handleOpenActionReport(item)}
                        />
                        <div>
                          <label className="flex items-center justify-between text-sm font-medium text-slate-700 mb-1">
                              このサブステップの添付資料
                              <button onClick={() => subStepAttachmentInputRef.current?.click()} className="p-1 hover:bg-slate-100 rounded-full" title="添付ファイルを追加">
                                  <PaperClipIcon className="w-4 h-4 text-slate-600"/>
                              </button>
                          </label>
                          {renderAttachments(selectedSubStep.attachments, (id) => handleRemoveAttachment(id, 'substep'))}
                          <input type="file" ref={subStepAttachmentInputRef} onChange={(e) => handleAttachmentChange(e, 'substep')} className="hidden"/>
                        </div>
                    </div>
                    ) : (
                    <div className="flex items-center justify-center h-full text-center text-slate-500 text-sm">
                        <p>左のフローチャートから<br/>サブステップを選択して詳細を編集します。</p>
                    </div>
                    )}
                </div>
            </section>
        </div>
      </main>
    </div>
  );
};

export default TaskDetailModal;
