package RecoveryPostSwap;

import java.util.Date;

public class SwapEntity {

	public String orgAircraft = "";
	public String swappedAricraft = "";
	public Date recoveryEnd = null;

	SwapEntity(String originalAc, String swappedAc) {

		this.orgAircraft = originalAc;
		this.swappedAricraft = swappedAc;
	}
}
