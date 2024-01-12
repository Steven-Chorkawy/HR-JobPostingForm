import { Shimmer } from '@fluentui/react';
import * as React from 'react';

export default class LoadingComponent extends React.Component<any, any> {
    public render(): React.ReactElement<any> {
        let style = { padding: "5px" };
        return (
            <div>
                <Shimmer style={style} />
                <Shimmer style={style} width="75%" />
                <Shimmer style={style} width="50%" />
                <Shimmer style={style} />
                <Shimmer style={style} width="75%" />
                <Shimmer style={style} width="50%" />
            </div>
        );
    }
}
