import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class PreambleTextField implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _container: HTMLDivElement;
    private _inputElement: HTMLInputElement;
    private _notifyOutputChanged: () => void;
    private _value: string;

    /**
     * Empty constructor.
     */
    constructor() {
        // Empty
    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        // Store references
        this._container = container;
        this._notifyOutputChanged = notifyOutputChanged;
        
        // Create the main container div
        const wrapperDiv = document.createElement("div");
        wrapperDiv.className = "preamble-textfield-container";
        
        // Create input element
        this._inputElement = document.createElement("input");
        this._inputElement.className = "ms-style-singleline";
        this._inputElement.type = "text";
        this._inputElement.value = "";
        this._inputElement.placeholder = "";
        
        // Add event listener for input changes
        this._inputElement.addEventListener("input", this.onInputChange.bind(this));
        
        // Append elements to container
        wrapperDiv.appendChild(this._inputElement);
        this._container.appendChild(wrapperDiv);
        
        // Initialize with current context
        this.updateView(context);
    }

    /**
     * Handle input change events
     */
    private onInputChange(event: Event): void {
        const target = event.target as HTMLInputElement;
        this._value = target.value;
        this._notifyOutputChanged();
    }

    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        // Set placeholder from the placeholderText property
        const placeholderValue = context.parameters.placeholder?.raw || "Enter text here...";
        this._inputElement.placeholder = placeholderValue;
        
        // Get the current value from the bound field
        const currentValue = context.parameters.ColumnInput?.raw || "";
        
        // Only update if the values are different to prevent infinite loops
        if (this._inputElement.value !== currentValue) {
            this._inputElement.value = currentValue;
            this._value = currentValue;
        }
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as "bound" or "output"
     */
    public getOutputs(): IOutputs {
        return {
            ColumnInput: this._value
        };
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        // Remove event listeners to prevent memory leaks
        if (this._inputElement) {
            this._inputElement.removeEventListener("input", this.onInputChange.bind(this));
        }
    }
}
