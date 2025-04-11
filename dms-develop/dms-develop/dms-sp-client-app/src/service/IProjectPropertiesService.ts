import { IProjectProperties } from "../interfaces/IProjectProperties";

export interface IProjectPropertiesService {
    getProperties(): Promise<IProjectProperties>;
    getProperty(propertyName: keyof IProjectProperties): Promise<string | undefined>;
    setProperty(propertyName: keyof IProjectProperties, value: string): Promise<void>;
    setProperties(properties: Partial<Record<keyof IProjectProperties, string>>): Promise<void>;
}