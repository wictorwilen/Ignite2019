import * as React from "react";
import { Text, Avatar, AvatarProps, Divider, Grid, Extendable, ShorthandValue } from "@stardust-ui/react";

export interface ICustomerCardProps {
  firstName?: string;
  lastName?: string;
  country?: string;
  email?: string;
  avatar?: ShorthandValue<AvatarProps>;
}

class CustomerCard extends React.Component<Extendable<ICustomerCardProps>, any> {
  public render() {
    const {
      firstName,
      lastName,
      country,
      email,
      avatar,
      ...restProps
    } = this.props;
    return (
      <Grid
        columns="80% 20%"
        styles={{ width: "320px", padding: "10px 20px 10px 10px"}}
        {...restProps}
      >
        <div>
          <Text size={"medium"} weight={"bold"} as="div">
            {firstName} {lastName}
          </Text>
          <Text muted as="div">
            {status}
          </Text>
          <Divider  />
          {country && (
            <Text muted as="div">
              {country}
            </Text>
          )}
          {email && (
            <Text muted as="div">
              {email}
            </Text>
          )}
        </div>
        {Avatar.create(avatar, {
          defaultProps: {
            size: "largest",
            name: `${firstName} ${lastName}`,
          },
        })}
      </Grid>
    );
  }
}

export default CustomerCard;
