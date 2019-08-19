import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as $ from "jquery";
import * as Bootstrap from "bootstrap";

export class CompositeAddress implements ComponentFramework.StandardControl<IInputs, IOutputs> {

    private notifyOutputChanged: () => void;
    private mainAddressField: HTMLInputElement;
    private fullAddressFormat: string;
    private currentValues: any = new Object();
    private currentLabels: any = new Object();
    private allControls: any = new Object();

    constructor() {

    }

    public init(context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement) {
        Bootstrap.hasOwnProperty("test");

        this.notifyOutputChanged = notifyOutputChanged;

        this.fullAddressFormat = context.parameters.format.raw;

        this.mainAddressField = document.createElement("input");
        this.mainAddressField.className = "compositeAddress";
        this.mainAddressField.readOnly = true;

        if (context.parameters.value.raw !== null) {
            this.mainAddressField.value = context.parameters.value.raw.replace(/\n/gm, ', ');
        } else {
            this.mainAddressField.value = "---";
        }

        this.mainAddressField.addEventListener("mouseenter", () => {
            this.mainAddressField.className = "compositeAddressFocused";
        });
        this.mainAddressField.addEventListener("mouseleave", () => {
            this.mainAddressField.className = "compositeAddress";
        });

        container.appendChild(this.mainAddressField);

        let popupDiv = document.createElement("div");

        this.addControltoPopup(context.parameters.street1, popupDiv, "street1");

        if (context.parameters.street2visible.raw === "yes") {
            this.addControltoPopup(context.parameters.street2, popupDiv, "street2");
        }

        if (context.parameters.street3visible.raw === "yes") {
            this.addControltoPopup(context.parameters.street3, popupDiv, "street3");
        }

        this.addControltoPopup(context.parameters.city, popupDiv, "city");

        if (context.parameters.countyvisible.raw === "yes") {
            this.addControltoPopup(context.parameters.county, popupDiv, "county");
        }

        if (context.parameters.statevisible.raw === "yes") {
            this.addControltoPopup(context.parameters.state, popupDiv, "state");
        }

        this.addControltoPopup(context.parameters.zipcode, popupDiv, "zipcode");

        if (context.parameters.countryvisible.raw === "yes") {
            this.addControltoPopup(context.parameters.country, popupDiv, "country");
        }

        $(this.mainAddressField).popover({
            content: popupDiv,
            html: true,
            placement: "bottom"
        });
    }

    private addControltoPopup(addressProperty: ComponentFramework.PropertyTypes.Property,
        container: HTMLDivElement,
        controlId: string): void {
        let rowDiv = document.createElement("div");
        rowDiv.className = "form-group";
        container.appendChild(rowDiv);

        if (addressProperty.type === "SingleLine.Text") {
            let stringAddressProperty = addressProperty as ComponentFramework.PropertyTypes.StringProperty;

            let fieldLabel = document.createElement("label");
            fieldLabel.innerText = stringAddressProperty.attributes === undefined ? "" : stringAddressProperty.attributes.DisplayName;
            rowDiv.appendChild(fieldLabel);

            let inputControl = document.createElement("input");
            inputControl.id = controlId;
            inputControl.className = "form-control d365StyleInput";
            inputControl.setAttribute("placeholder", "---");
            inputControl.value = stringAddressProperty.raw;
            inputControl.addEventListener("change", () => {
                let currentText = inputControl.value;

                this.currentValues[controlId] = currentText;
                this.currentLabels[controlId] = currentText === null ? "" : currentText;

                this.formatAddress();

                this.notifyOutputChanged();
            });

            this.currentValues[controlId] = stringAddressProperty.raw;
            this.currentLabels[controlId] = stringAddressProperty.formatted;
            this.allControls[controlId] = inputControl;

            rowDiv.appendChild(inputControl);
        } else if (addressProperty.type === "OptionSet") {
            let optionsetAddressProperty = addressProperty as ComponentFramework.PropertyTypes.OptionSetProperty;

            let fieldLabel = document.createElement("label");
            fieldLabel.innerText = optionsetAddressProperty.attributes === undefined
                ? ""
                : optionsetAddressProperty.attributes.DisplayName;
            rowDiv.appendChild(fieldLabel);

            let selectControl = document.createElement("select");
            selectControl.id = controlId;
            selectControl.className = "form-control d365StyleInput";

            let option: HTMLOptionElement = document.createElement("option");
            option.innerHTML = "---Select---";
            selectControl.add(option);

            if (optionsetAddressProperty.attributes !== undefined) {
                optionsetAddressProperty.attributes.Options.forEach(optionRecord => {
                    option = document.createElement("option");
                    option.innerHTML = optionRecord.Label;
                    option.value = optionRecord.Value.toString();

                    if (optionsetAddressProperty.raw === optionRecord.Value) {
                        option.selected = true;
                    }

                    selectControl.add(option);
                });
            }

            selectControl.addEventListener("change", () => {
                if (selectControl.value === "") {
                    this.currentValues[controlId] = null;
                    this.currentLabels[controlId] = "";
                } else {
                    this.currentValues[controlId] = parseInt(selectControl.value);
                    this.currentLabels[controlId] = selectControl.options[selectControl.selectedIndex].label;
                }

                this.formatAddress();

                this.notifyOutputChanged();
            });

            this.currentValues[controlId] = optionsetAddressProperty.raw;
            this.currentLabels[controlId] = optionsetAddressProperty.formatted;
            this.allControls[controlId] = selectControl;

            rowDiv.appendChild(selectControl);
        } else {
            throw new Error(`Composite Address: can't render control for ${addressProperty.type} type of field`);
        }
    }

    private formatAddress(): void {
        let formattedAddress = this.fullAddressFormat;

        formattedAddress = formattedAddress.replace(/\\n/gm, '\n');

        for (let propName in this.currentLabels) {
            let propertyValue = this.currentLabels[propName] === null ? "" : this.currentLabels[propName];

            formattedAddress = formattedAddress.replace(`{${propName}}`, propertyValue);
        }

        this.currentValues.value = formattedAddress;
        this.mainAddressField.value = formattedAddress.replace(/\n/gm, ', ');
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        let formatChanged: boolean = false;

        context.updatedProperties.forEach(updatedProperty => {
            if (!context.parameters.hasOwnProperty(updatedProperty)) {
                return ;
            }

            formatChanged = true;

            let updatedPropertyObject = (context.parameters as any)[updatedProperty];

            if (this.allControls.hasOwnProperty(updatedProperty)) {
                $(this.allControls[updatedProperty]).val(updatedPropertyObject.raw);
            }

            this.currentValues[updatedProperty] = updatedPropertyObject.raw;
            this.currentLabels[updatedProperty] = updatedPropertyObject.formatted;
        });

        if (formatChanged) {
            this.formatAddress();
            this.notifyOutputChanged();
        }
    }

    public getOutputs(): IOutputs {
        return this.currentValues as IOutputs;
    }

    public destroy(): void {
        // Add code to cleanup control if necessary
    }
}