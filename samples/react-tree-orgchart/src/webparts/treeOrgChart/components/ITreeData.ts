
export interface ITreeData {
   // eslint-disable-next-line @typescript-eslint/no-explicit-any
   title: any;
   expanded ?: boolean;
   children ? : ITreeData[]|null;
}
